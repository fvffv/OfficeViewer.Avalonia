using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Avalonia;
using global::Avalonia.Controls;
using global::Avalonia.Controls.Primitives;
using global::Avalonia.Controls.Shapes;
using global::Avalonia.Input;
using global::Avalonia.Interactivity;
using global::Avalonia.Layout;
using global::Avalonia.Media;
using global::Avalonia.Media.Imaging;
using global::Avalonia.Platform;
using global::Avalonia.Threading;

namespace OfficeViewer.Avalonia.Pdf;

/// <summary>
/// Pure managed, continuous PDF viewer. The component mirrors vue-office's
/// scroll virtualization: only visible pages and adjacent buffers materialize
/// Avalonia text/image controls; moving out of the cache disposes bitmaps.
/// </summary>
public sealed class PdfViewer : UserControl, IDisposable
{
    private const double PageSpacing = 18;
    private readonly ScrollViewer _scrollViewer;
    private readonly StackPanel _pages;
    private readonly LayoutTransformControl _zoomHost;
    private readonly ScaleTransform _zoomTransform = new();
    private readonly Dictionary<int, PageHost> _pageHosts = [];
    private readonly Dictionary<int, List<IDisposable>> _ownedBitmaps = [];
    private PdfDocumentModel? _document;
    private List<IDisposable>? _renderingBitmaps;
    private CancellationTokenSource? _loadCancellation;
    private Task? _documentLoadTask;
    private long _generation;
    private double _zoom = 1;
    private bool _settingDocument;
    private bool _disposed;

    /// <summary>Gets the number of PDF pages after a successful load.</summary>
    public static readonly StyledProperty<int> PageCountProperty =
        AvaloniaProperty.Register<PdfViewer, int>(nameof(PageCount));

    public int PageCount
    {
        get => GetValue(PageCountProperty);
        private set => SetValue(PageCountProperty, value);
    }

    /// <summary>
    /// Gets or sets the PDF bytes. Assign a new byte array to load it; assign
    /// <see langword="null"/> to cancel loading and release viewer resources.
    /// </summary>
    public static readonly StyledProperty<byte[]?> DocumentProperty =
        AvaloniaProperty.Register<PdfViewer, byte[]?>(nameof(Document));

    public byte[]? Document
    {
        get => GetValue(DocumentProperty);
        set => SetValue(DocumentProperty, value);
    }

    public PdfViewer()
    {
        _pages = new StackPanel
        {
            Spacing = PageSpacing,
            Margin = new Thickness(24),
            HorizontalAlignment = HorizontalAlignment.Center
        };
        _zoomHost = new LayoutTransformControl
        {
            Child = _pages,
            LayoutTransform = _zoomTransform,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Top
        };
        _scrollViewer = new ScrollViewer
        {
            Content = _zoomHost,
            HorizontalContentAlignment = HorizontalAlignment.Center,
            VerticalContentAlignment = VerticalAlignment.Top,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        };
        _scrollViewer.ScrollChanged += OnScrollChanged;
        AddHandler(InputElement.PointerWheelChangedEvent, OnPointerWheelChanged, RoutingStrategies.Tunnel);
        Content = _scrollViewer;
    }

    public async Task LoadAsync(string path, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        await LoadAsync(await File.ReadAllBytesAsync(path, cancellationToken), cancellationToken);
    }

    /// <summary>Loads PDF bytes and updates <see cref="Document"/>.</summary>
    public Task LoadAsync(byte[] document, CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentNullException.ThrowIfNull(document);
        _settingDocument = true;
        try { SetCurrentValue(DocumentProperty, document); }
        finally { _settingDocument = false; }
        return _documentLoadTask = LoadDocumentAsync(document, cancellationToken);
    }

    private async Task LoadDocumentAsync(byte[] document, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var generation = Interlocked.Increment(ref _generation);
        var requestCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        Interlocked.Exchange(ref _loadCancellation, requestCancellation)?.Cancel();
        PdfDocumentModel? parsed = null;
        await Dispatcher.UIThread.InvokeAsync(() =>
        {
            if (generation == Volatile.Read(ref _generation)) ReleaseDocument();
        });
        try
        {
            parsed = await Task.Run(() => new PdfParser().Parse(document), requestCancellation.Token).ConfigureAwait(false);
            requestCancellation.Token.ThrowIfCancellationRequested();
            if (generation != Volatile.Read(ref _generation)) return;
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                if (generation != Volatile.Read(ref _generation)) return;
                Render(parsed);
                parsed = null; // ownership transferred to the viewer
            });
        }
        catch (OperationCanceledException) when (requestCancellation.IsCancellationRequested)
        {
            // A newer document owns the viewer now.
        }
        catch (Exception error)
        {
            if (generation != Volatile.Read(ref _generation)) return;
            await Dispatcher.UIThread.InvokeAsync(() => ShowError(error));
        }
        finally
        {
            parsed?.Dispose();
            Interlocked.CompareExchange(ref _loadCancellation, null, requestCancellation);
            requestCancellation.Dispose();
        }
    }

    public void Clear()
    {
        if (Document is not null)
        {
            SetCurrentValue(DocumentProperty, null);
            return;
        }
        ClearDocument();
    }

    private void ClearDocument()
    {
        Interlocked.Increment(ref _generation);
        Interlocked.Exchange(ref _loadCancellation, null)?.Cancel();
        ReleaseDocument();
    }

    private void ReleaseDocument()
    {
        ClearVisuals();
        _document?.Dispose();
        _document = null;
    }

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property != DocumentProperty || _settingDocument) return;
        var document = change.GetNewValue<byte[]?>();
        if (document is null)
        {
            _documentLoadTask = null;
            ClearDocument();
        }
        else if (!_disposed)
        {
            _documentLoadTask = LoadDocumentAsync(document, CancellationToken.None);
        }
    }

    private void Render(PdfDocumentModel document)
    {
        ReleaseDocument();
        _document = document;
        for (var index = 0; index < document.PageCount; index++)
        {
            var page = document.GetPageInfo(index);
            var slot = new Border
            {
                Width = page.Width,
                Height = page.Height,
                Background = Brushes.White,
                BorderBrush = Brush.Parse("#D0D5DB"),
                BorderThickness = new Thickness(1),
                ClipToBounds = true,
                BoxShadow = new BoxShadows(new BoxShadow { Blur = 10, OffsetX = 0, OffsetY = 2, Color = Color.FromArgb(48, 0, 0, 0) })
            };
            _pageHosts[index] = new PageHost(slot);
            _pages.Children.Add(slot);
        }
        PageCount = document.PageCount;
        _scrollViewer.Offset = default;
        UpdateVirtualizedPages();
    }

    private void OnScrollChanged(object? sender, ScrollChangedEventArgs e) => UpdateVirtualizedPages();

    private void UpdateVirtualizedPages()
    {
        if (_document is null || _pageHosts.Count == 0) return;
        var offset = _scrollViewer.Offset.Y / _zoom;
        var viewport = Math.Max(1, _scrollViewer.Viewport.Height / _zoom);
        var firstVisible = FindPageAt(offset);
        var lastVisible = FindPageAt(offset + viewport);
        var firstKept = Math.Max(0, firstVisible - 2);
        var lastKept = Math.Min(_document.PageCount - 1, lastVisible + 2);
        foreach (var index in _pageHosts.Keys.ToArray())
        {
            if (index >= firstKept && index <= lastKept) MaterializePage(index);
            else ReleasePage(index);
        }
    }

    private int FindPageAt(double position)
    {
        if (_document is null) return 0;
        var top = _pages.Margin.Top;
        for (var index = 0; index < _document.PageCount; index++)
        {
            top += _document.GetPageInfo(index).Height;
            if (position <= top) return index;
            top += PageSpacing;
        }
        return _document.PageCount - 1;
    }

    private void MaterializePage(int index)
    {
        if (_document is null || !_pageHosts.TryGetValue(index, out var host) || host.Canvas is not null) return;
        var page = _document.GetPage(index);
        var canvas = new Canvas { Width = page.Width, Height = page.Height, ClipToBounds = true };

        var bitmaps = new List<IDisposable>();
        _renderingBitmaps = bitmaps;
        try
        {
            // WPS places its trial mark at the end of the content stream. It is
            // marked as a PDF /Artifact watermark, so materialize it first and
            // let the document body remain readable above it.
            foreach (var watermark in page.DrawOperations.OfType<PdfPageImageDraw>().Where(item => item.Image.IsWatermark))
                AddImage(canvas, watermark.Image);
            foreach (var operation in page.DrawOperations)
            {
                switch (operation)
                {
                    case PdfPageFillDraw fill: AddFill(canvas, fill.Fill); break;
                    case PdfPageLineDraw line: AddLine(canvas, page, line.Line); break;
                    case PdfPageTextDraw text: AddText(canvas, page, text.Text); break;
                    case PdfPageImageDraw image when !image.Image.IsWatermark: AddImage(canvas, image.Image); break;
                }
            }
        }
        finally
        {
            _renderingBitmaps = null;
        }

        host.Canvas = canvas;
        host.Slot.Child = canvas;
        _ownedBitmaps[index] = bitmaps;
    }

    private static void AddLine(Canvas canvas, PdfPageModel page, PdfLineSegment line)
    {
        var control = new Line
        {
            StartPoint = new Point(line.StartX, page.Height - line.StartY),
            EndPoint = new Point(line.EndX, page.Height - line.EndY),
            Stroke = new SolidColorBrush(line.Color),
            StrokeThickness = Math.Max(.5, line.Thickness)
        };
        canvas.Children.Add(control);
    }

    private static void AddFill(Canvas canvas, PdfFilledPolygon fill)
    {
        if (fill.Points.Count < 3) return;
        canvas.Children.Add(new Polygon
        {
            Points = new Points(fill.Points.Select(point => new Point(point.X, point.Y))),
            Fill = new SolidColorBrush(fill.Color)
        });
    }

    private static void AddText(Canvas canvas, PdfPageModel page, PdfTextSegment segment)
    {
        var text = new TextBlock
        {
            Text = segment.Text,
            Foreground = new SolidColorBrush(segment.Color),
            FontSize = Math.Max(1, segment.FontSize),
            TextWrapping = TextWrapping.NoWrap
        };
        if (!string.IsNullOrWhiteSpace(segment.FontFamily)) text.FontFamily = new FontFamily(segment.FontFamily);
        Canvas.SetLeft(text, segment.Left);
        // Avalonia positions text at its layout top, PDF positions it at the baseline.
        // The 0.82 ascent is a stable approximation for standard Latin/CJK fonts.
        Canvas.SetTop(text, Math.Max(0, page.Height - segment.Baseline - segment.FontSize * .82));
        canvas.Children.Add(text);
    }

    private void AddImage(Canvas canvas, PdfImageModel image)
    {
        try
        {
            var decoded = DecodeImage(image.Payload);
            if (decoded is null) return;
            _renderingBitmaps?.Add(decoded.Value.Owner);
            var control = new Image
            {
                Source = decoded.Value.Source,
                Width = image.Width,
                Height = image.Height,
                Stretch = Stretch.Fill,
                Opacity = Math.Clamp(image.Opacity, 0, 1)
            };
            Canvas.SetLeft(control, image.Left);
            Canvas.SetTop(control, image.Top);
            canvas.Children.Add(control);
        }
        catch
        {
            // A damaged or unsupported embedded image must not prevent text from rendering.
        }
    }

    private static (IImage Source, IDisposable Owner)? DecodeImage(PdfImagePayload payload)
    {
        // JPEG/JPX streams remain encoded and Avalonia's normal bitmap decoder can
        // load them directly. A JPEG can still carry a separate /SMask; WPS uses
        // that combination for trial watermarks, so merge that alpha channel first.
        if (payload.Filters.Any(filter => filter is "DCTDecode" or "DCT" or "JPXDecode"))
        {
            var bitmap = DecodeCompressedBitmap(payload.EncodedBytes);
            if (bitmap is null) return null;
            var mask = DecodeSoftMaskBitmap(payload);
            if (mask is null) return (bitmap, bitmap);
            try
            {
                var merged = ApplySoftMask(bitmap, mask);
                if (merged is null) return (bitmap, bitmap);
                bitmap.Dispose();
                return (merged, merged);
            }
            finally
            {
                mask.Dispose();
            }
        }
        if (!payload.Filters.All(filter => filter is "FlateDecode" or "Fl")) return null;
        if (payload.BitsPerComponent != 8 || payload.PixelWidth <= 0 || payload.PixelHeight <= 0) return null;
        var channels = string.Equals(payload.ColorSpace, "DeviceGray", StringComparison.Ordinal) ? 1 : 3;
        if (!string.Equals(payload.ColorSpace, "DeviceGray", StringComparison.Ordinal) && !string.Equals(payload.ColorSpace, "DeviceRGB", StringComparison.Ordinal)) return null;
        var source = Inflate(payload.EncodedBytes);
        var pixelCount = checked(payload.PixelWidth * payload.PixelHeight);
        if (source.Length < pixelCount * channels) return null;
        var alpha = payload.SoftMaskBytes is null ? null : Inflate(payload.SoftMaskBytes);
        var pixels = new byte[checked(pixelCount * 4)];
        for (var index = 0; index < pixelCount; index++)
        {
            var sourceIndex = index * channels;
            var red = channels == 1 ? source[sourceIndex] : source[sourceIndex];
            var green = channels == 1 ? red : source[sourceIndex + 1];
            var blue = channels == 1 ? red : source[sourceIndex + 2];
            var opacity = alpha is not null && index < alpha.Length ? alpha[index] : (byte)255;
            if (payload.ColorKeyMask is { Length: 6 } key &&
                red >= key[0] && red <= key[1] && green >= key[2] && green <= key[3] && blue >= key[4] && blue <= key[5]) opacity = 0;
            var target = index * 4;
            pixels[target] = (byte)(blue * opacity / 255);
            pixels[target + 1] = (byte)(green * opacity / 255);
            pixels[target + 2] = (byte)(red * opacity / 255);
            pixels[target + 3] = opacity;
        }
        var output = new WriteableBitmap(new PixelSize(payload.PixelWidth, payload.PixelHeight), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);
        using (var frame = output.Lock())
        {
            var rowBytes = payload.PixelWidth * 4;
            for (var row = 0; row < payload.PixelHeight; row++)
                Marshal.Copy(pixels, row * rowBytes, IntPtr.Add(frame.Address, row * frame.RowBytes), rowBytes);
        }
        return (output, output);
    }

    private static Bitmap? DecodeCompressedBitmap(byte[] bytes)
    {
        try
        {
            using var stream = new MemoryStream(bytes, writable: false);
            return new Bitmap(stream);
        }
        catch
        {
            return null;
        }
    }

    private static Bitmap? DecodeSoftMaskBitmap(PdfImagePayload payload)
    {
        if (payload.SoftMaskBytes is not { Length: > 0 }) return null;
        var filters = payload.SoftMaskFilters ?? [];
        return filters.Any(filter => filter is "DCTDecode" or "DCT" or "JPXDecode")
            ? DecodeCompressedBitmap(payload.SoftMaskBytes)
            : null;
    }

    private static WriteableBitmap? ApplySoftMask(Bitmap source, Bitmap mask)
    {
        if (source.PixelSize != mask.PixelSize || source.PixelSize.Width <= 0 || source.PixelSize.Height <= 0) return null;
        var size = source.PixelSize;
        var color = new WriteableBitmap(size, new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);
        var alpha = new WriteableBitmap(size, new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);
        try
        {
            using (var initialColorFrame = color.Lock()) source.CopyPixels(initialColorFrame);
            using (var initialAlphaFrame = alpha.Lock()) mask.CopyPixels(initialAlphaFrame);
            using var colorFrame = color.Lock();
            using var alphaFrame = alpha.Lock();
            var colorRow = new byte[size.Width * 4];
            var alphaRow = new byte[size.Width * 4];
            for (var row = 0; row < size.Height; row++)
            {
                Marshal.Copy(IntPtr.Add(colorFrame.Address, row * colorFrame.RowBytes), colorRow, 0, colorRow.Length);
                Marshal.Copy(IntPtr.Add(alphaFrame.Address, row * alphaFrame.RowBytes), alphaRow, 0, alphaRow.Length);
                for (var column = 0; column < size.Width; column++)
                {
                    var offset = column * 4;
                    var opacity = alphaRow[offset + 2]; // Gray mask decoded as BGRA.
                    colorRow[offset] = (byte)(colorRow[offset] * opacity / 255);
                    colorRow[offset + 1] = (byte)(colorRow[offset + 1] * opacity / 255);
                    colorRow[offset + 2] = (byte)(colorRow[offset + 2] * opacity / 255);
                    colorRow[offset + 3] = (byte)(colorRow[offset + 3] * opacity / 255);
                }
                Marshal.Copy(colorRow, 0, IntPtr.Add(colorFrame.Address, row * colorFrame.RowBytes), colorRow.Length);
            }
            return color;
        }
        catch
        {
            color.Dispose();
            return null;
        }
        finally
        {
            alpha.Dispose();
        }
    }

    private static byte[] Inflate(byte[] data)
    {
        try { return Inflate(data, stream => new ZLibStream(stream, CompressionMode.Decompress)); }
        catch (InvalidDataException) { return Inflate(data, stream => new DeflateStream(stream, CompressionMode.Decompress)); }
    }

    private static byte[] Inflate(byte[] data, Func<Stream, Stream> create)
    {
        using var source = new MemoryStream(data, writable: false);
        using var compressed = create(source);
        using var output = new MemoryStream();
        compressed.CopyTo(output);
        return output.ToArray();
    }

    private void ReleasePage(int index)
    {
        if (!_pageHosts.TryGetValue(index, out var host) || host.Canvas is null) return;
        host.Slot.Child = null;
        host.Canvas = null;
        if (_ownedBitmaps.Remove(index, out var bitmaps))
            foreach (var bitmap in bitmaps) bitmap.Dispose();
    }

    private void OnPointerWheelChanged(object? sender, PointerWheelEventArgs e)
    {
        if (!e.KeyModifiers.HasFlag(KeyModifiers.Control)) return;
        _zoom = Math.Clamp(_zoom * (e.Delta.Y > 0 ? 1.1 : 1 / 1.1), .5, 2.5);
        _zoomTransform.ScaleX = _zoom;
        _zoomTransform.ScaleY = _zoom;
        _scrollViewer.Offset = new Vector(0, _scrollViewer.Offset.Y);
        UpdateVirtualizedPages();
        e.Handled = true;
    }

    private void ShowError(Exception error)
    {
        ClearDocument();
        _pages.Children.Add(new Border
        {
            Background = Brush.Parse("#FFF1F2"),
            BorderBrush = Brush.Parse("#FDA4AF"),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(16),
            Child = new TextBlock { Text = $"无法加载 PDF：{error.Message}", Foreground = Brush.Parse("#9F1239"), TextWrapping = TextWrapping.Wrap }
        });
    }

    private void ClearVisuals()
    {
        foreach (var index in _pageHosts.Keys.ToArray()) ReleasePage(index);
        _ownedBitmaps.Clear();
        _pageHosts.Clear();
        _pages.Children.Clear();
        PageCount = 0;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Clear();
    }

    private sealed class PageHost(Border slot)
    {
        public Border Slot { get; } = slot;
        public Canvas? Canvas { get; set; }
    }
}
