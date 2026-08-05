using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Avalonia;
using global::Avalonia.Controls;
using global::Avalonia.Controls.Documents;
using global::Avalonia.Controls.Primitives;
using global::Avalonia.Controls.Shapes;
using global::Avalonia.Input;
using global::Avalonia.Interactivity;
using global::Avalonia.Layout;
using global::Avalonia.Media;
using global::Avalonia.Media.Imaging;
using global::Avalonia.Threading;

namespace OfficeViewer.Avalonia.Pptx;

/// <summary>
/// Native, continuous PresentationML preview.  Slide shells keep the document's
/// natural vertical flow, while pictures are decoded only for the visible slides
/// and their immediate neighbours.
/// </summary>
public sealed class PptxViewer : UserControl, IDisposable
{
    private const double SlideSpacing = 18;
    private readonly ScrollViewer _scrollViewer;
    private readonly StackPanel _slides;
    private readonly LayoutTransformControl _zoomHost;
    private readonly ScaleTransform _zoomTransform = new();
    private readonly Dictionary<int, SlideHost> _slideHosts = [];
    private readonly Dictionary<int, List<IDisposable>> _ownedBitmaps = [];
    private PptxDocumentModel? _document;
    private List<IDisposable>? _renderingBitmaps;
    private CancellationTokenSource? _loadCancellation;
    private Task? _documentLoadTask;
    private long _generation;
    private double _zoom = 1;
    private bool _settingDocument;
    private bool _disposed;

    /// <summary>Gets the number of slides after a successful presentation load.</summary>
    public static readonly StyledProperty<int> PageCountProperty =
        AvaloniaProperty.Register<PptxViewer, int>(nameof(PageCount));

    public int PageCount
    {
        get => GetValue(PageCountProperty);
        private set => SetValue(PageCountProperty, value);
    }

    /// <summary>
    /// Gets or sets the PPTX package bytes. Assign a new byte array to load it;
    /// assign <see langword="null"/> to cancel loading and release viewer resources.
    /// </summary>
    public static readonly StyledProperty<byte[]?> DocumentProperty =
        AvaloniaProperty.Register<PptxViewer, byte[]?>(nameof(Document));

    public byte[]? Document
    {
        get => GetValue(DocumentProperty);
        set => SetValue(DocumentProperty, value);
    }

    public PptxViewer()
    {
        _slides = new StackPanel
        {
            Spacing = SlideSpacing,
            Margin = new Thickness(24),
            HorizontalAlignment = HorizontalAlignment.Center
        };
        _zoomHost = new LayoutTransformControl
        {
            Child = _slides,
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
            // Intentionally no Background: embedding applications own the surface.
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

    /// <summary>Loads a PPTX package and updates <see cref="Document"/>.</summary>
    public Task LoadAsync(byte[] document, CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentNullException.ThrowIfNull(document);
        _settingDocument = true;
        try { SetCurrentValue(DocumentProperty, document); }
        finally { _settingDocument = false; }
        return _documentLoadTask = LoadDocumentAsync(document, cancellationToken);
    }

    /// <summary>Loads a PPTX stream without taking ownership of the supplied stream.</summary>
    public async Task LoadAsync(Stream source, CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var document = await global::OfficeViewer.Avalonia.DocumentData.ReadAsync(source, cancellationToken).ConfigureAwait(false);
        await LoadAsync(document, cancellationToken).ConfigureAwait(false);
    }

    private async Task LoadDocumentAsync(byte[] document, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var generation = Interlocked.Increment(ref _generation);
        var requestCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        Interlocked.Exchange(ref _loadCancellation, requestCancellation)?.Cancel();
        await Dispatcher.UIThread.InvokeAsync(ReleaseDocument);
        try
        {
            var parsed = await Task.Run(() => new PptxParser().Parse(document), requestCancellation.Token).ConfigureAwait(false);
            requestCancellation.Token.ThrowIfCancellationRequested();
            if (generation != Volatile.Read(ref _generation)) return;
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                if (generation == Volatile.Read(ref _generation)) Render(parsed);
            });
        }
        catch (OperationCanceledException) when (requestCancellation.IsCancellationRequested)
        {
            // A newer document owns the viewer.
        }
        catch (Exception error)
        {
            if (generation != Volatile.Read(ref _generation)) return;
            await Dispatcher.UIThread.InvokeAsync(() => ShowError(error));
        }
        finally
        {
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

    private void Render(PptxDocumentModel document)
    {
        ReleaseDocument();
        _document = document;
        for (var index = 0; index < document.Slides.Count; index++)
        {
            // This fixed-size layout reservation stays white while its visual canvas
            // is materialized lazily, so slides without an explicit background do
            // not become transparent and a 100-slide deck still avoids eager media decoding.
            var slot = new Border
            {
                Width = document.Width,
                Height = document.Height,
                Background = Brushes.White,
                ClipToBounds = true
            };
            _slideHosts[index] = new SlideHost(slot);
            _slides.Children.Add(slot);
        }
        PageCount = document.Slides.Count;
        _scrollViewer.Offset = default;
        UpdateVirtualizedSlides();
    }

    private void OnScrollChanged(object? sender, ScrollChangedEventArgs e) => UpdateVirtualizedSlides();

    private void UpdateVirtualizedSlides()
    {
        if (_document is null || _slideHosts.Count == 0) return;
        var unscaledOffset = _scrollViewer.Offset.Y / _zoom;
        var unscaledViewport = Math.Max(1, _scrollViewer.Viewport.Height / _zoom);
        var slideFootprint = _document.Height + SlideSpacing;
        var firstVisible = Math.Clamp((int)Math.Floor((unscaledOffset - _slides.Margin.Top) / slideFootprint), 0, _document.Slides.Count - 1);
        var visibleCount = Math.Max(1, (int)Math.Ceiling(unscaledViewport / slideFootprint) + 1);
        var firstKept = Math.Max(0, firstVisible - 1);
        var lastKept = Math.Min(_document.Slides.Count - 1, firstVisible + visibleCount);

        foreach (var index in _slideHosts.Keys.ToArray())
        {
            if (index >= firstKept && index <= lastKept) MaterializeSlide(index);
            else ReleaseSlide(index);
        }
    }

    private void MaterializeSlide(int index)
    {
        if (_document is null || !_slideHosts.TryGetValue(index, out var host) || host.Canvas is not null) return;
        var canvas = new Canvas
        {
            Width = _document.Width,
            Height = _document.Height,
            ClipToBounds = true,
            Background = ToBrush(_document.Slides[index].Background) ?? Brushes.White
        };
        var bitmaps = new List<IDisposable>();
        _renderingBitmaps = bitmaps;
        try
        {
            foreach (var element in _document.Slides[index].Elements) AddElement(canvas, element);
        }
        finally
        {
            _renderingBitmaps = null;
        }
        host.Canvas = canvas;
        host.Slot.Child = canvas;
        _ownedBitmaps[index] = bitmaps;
    }

    private void ReleaseSlide(int index)
    {
        if (!_slideHosts.TryGetValue(index, out var host) || host.Canvas is null) return;
        // Detach controls before disposing sources; this avoids a frame attempting to
        // paint a just-disposed Skia bitmap during a fast scroll.
        host.Slot.Child = null;
        host.Canvas = null;
        if (_ownedBitmaps.Remove(index, out var bitmaps))
            foreach (var bitmap in bitmaps) bitmap.Dispose();
    }

    private void AddElement(Canvas canvas, PptxElement element)
    {
        Control? control = element switch
        {
            PptxShape shape => CreateShape(shape),
            PptxPicture picture => CreatePicture(picture),
            PptxLine line => CreateLine(line),
            PptxChart chart => CreateChart(chart),
            _ => null
        };
        if (control is null) return;
        control.Width = Math.Max(0, element.Width);
        control.Height = Math.Max(0, element.Height);
        if (element.Rotation != 0 || element.FlipHorizontal || element.FlipVertical)
        {
            var transforms = new TransformGroup();
            if (element.FlipHorizontal || element.FlipVertical)
                transforms.Children.Add(new ScaleTransform(element.FlipHorizontal ? -1 : 1, element.FlipVertical ? -1 : 1));
            if (element.Rotation != 0) transforms.Children.Add(new RotateTransform(element.Rotation));
            control.RenderTransform = transforms;
            control.RenderTransformOrigin = RelativePoint.Center;
        }
        Canvas.SetLeft(control, element.Left);
        Canvas.SetTop(control, element.Top);
        canvas.Children.Add(control);
    }

    private Control CreateShape(PptxShape shape)
    {
        var fill = ToBrush(shape.Fill);
        var stroke = ToBrush(shape.Stroke);
        var text = CreateText(shape);
        if (shape.Geometry == "ellipse")
        {
            var host = new Grid();
            host.Children.Add(new Ellipse { Fill = fill, Stroke = stroke, StrokeThickness = shape.StrokeThickness });
            if (text is not null) host.Children.Add(text);
            return host;
        }
        if (shape.Geometry is "triangle" or "homePlate" or "hexagon")
        {
            var host = new Grid();
            host.Children.Add(new Polygon { Points = GeometryPoints(shape.Geometry), Stretch = Stretch.Fill, Fill = fill, Stroke = stroke, StrokeThickness = shape.StrokeThickness });
            if (text is not null) host.Children.Add(text);
            return host;
        }
        return new Border
        {
            Background = fill,
            BorderBrush = stroke,
            BorderThickness = new Thickness(shape.StrokeThickness),
            CornerRadius = shape.Geometry is "roundRect" or "round1Rect" ? new CornerRadius(12) : default,
            Child = text
        };
    }

    private static TextBlock? CreateText(PptxShape shape)
    {
        if (shape.Paragraphs.Count == 0) return null;
        var text = new TextBlock
        {
            TextWrapping = TextWrapping.Wrap,
            Padding = shape.TextInsets,
            TextAlignment = shape.TextAlignment,
            VerticalAlignment = shape.VerticalAlignment,
            Foreground = Brushes.Black
        };
        for (var paragraphIndex = 0; paragraphIndex < shape.Paragraphs.Count; paragraphIndex++)
        {
            var paragraph = shape.Paragraphs[paragraphIndex];
            if (paragraphIndex > 0) text.Inlines!.Add(new LineBreak());
            if (paragraph.Alignment is { } alignment) text.TextAlignment = alignment;
            foreach (var item in paragraph.Runs)
            {
                var run = new Run(item.Text);
                if (!string.IsNullOrWhiteSpace(item.FontFamily)) run.FontFamily = new FontFamily(item.FontFamily);
                if (item.FontSize is > 0) run.FontSize = item.FontSize.Value;
                if (item.FontWeight is { } weight) run.FontWeight = weight;
                if (item.FontStyle is { } style) run.FontStyle = style;
                run.Foreground = ToBrush(item.Foreground) ?? Brushes.Black;
                if (item.Underline) run.TextDecorations = TextDecorations.Underline;
                text.Inlines!.Add(run);
            }
        }
        return text;
    }

    private Image? CreatePicture(PptxPicture picture)
    {
        if (_document is null) return null;
        try
        {
            using var package = _document.SourceBytes is { } sourceBytes
                ? new ZipArchive(new MemoryStream(sourceBytes, writable: false), ZipArchiveMode.Read)
                : ZipFile.OpenRead(_document.SourcePath ?? throw new InvalidOperationException("The presentation source is unavailable."));
            var entry = package.GetEntry(picture.PackagePath);
            if (entry is null) return null;
            using var compressed = entry.Open();
            using var stream = new MemoryStream((int)Math.Min(entry.Length, int.MaxValue));
            compressed.CopyTo(stream);
            stream.Position = 0;
            var bitmap = Bitmap.DecodeToWidth(stream, Math.Clamp((int)Math.Ceiling(picture.Width), 1, 1280), BitmapInterpolationMode.HighQuality);
            _renderingBitmaps?.Add(bitmap);
            return new Image { Source = bitmap, Stretch = Stretch.Fill };
        }
        catch (Exception) { return null; } // SVG/WMF or a damaged image must not fail the document.
    }

    private static Line CreateLine(PptxLine line) => new()
    {
        StartPoint = new Point(line.FlipHorizontal ? line.Width : 0, line.FlipVertical ? line.Height : 0),
        EndPoint = new Point(line.FlipHorizontal ? 0 : line.Width, line.FlipVertical ? 0 : line.Height),
        Stroke = ToBrush(line.Stroke),
        StrokeThickness = Math.Max(1, line.StrokeThickness)
    };

    private static Points GeometryPoints(string geometry) => geometry switch
    {
        "triangle" => new Points { new Point(.5, 0), new Point(1, 1), new Point(0, 1) },
        "homePlate" => new Points { new Point(0, 0), new Point(.84, 0), new Point(1, .5), new Point(.84, 1), new Point(0, 1) },
        "hexagon" => new Points { new Point(.25, 0), new Point(.75, 0), new Point(1, .5), new Point(.75, 1), new Point(.25, 1), new Point(0, .5) },
        _ => []
    };

    private static Canvas CreateChart(PptxChart chart)
    {
        var canvas = new Canvas { Width = chart.Width, Height = chart.Height, ClipToBounds = false };
        var width = Math.Max(chart.Width, 1);
        var height = Math.Max(chart.Height, 1);
        var left = Math.Min(46, width * .18);
        var top = Math.Min(30, height * .12);
        var right = Math.Min(12, width * .05);
        var bottom = Math.Min(86, height * .30);
        var plotWidth = Math.Max(1, width - left - right);
        var plotHeight = Math.Max(1, height - top - bottom);
        var maximum = Math.Max(1, Math.Ceiling(chart.Values.Max() * 2d) / 2d + .5d);
        const int gridCount = 9;
        for (var i = 0; i <= gridCount; i++)
        {
            var value = maximum * i / gridCount;
            var y = top + plotHeight * (1 - i / (double)gridCount);
            canvas.Children.Add(new Line { StartPoint = new Point(left, y), EndPoint = new Point(left + plotWidth, y), Stroke = Brush.Parse("#D9FFFFFF"), StrokeThickness = .8 });
            var label = new TextBlock { Text = value.ToString("0.#"), FontSize = 9, Foreground = Brushes.White, Width = Math.Max(1, left - 5), TextAlignment = TextAlignment.Right };
            Canvas.SetLeft(label, 0); Canvas.SetTop(label, y - 7); canvas.Children.Add(label);
        }
        var count = Math.Min(chart.Values.Count, Math.Max(1, chart.Categories.Count));
        var step = plotWidth / count;
        var barWidth = Math.Max(2, step * .35);
        var barBrush = ToBrush(chart.BarFill) ?? Brushes.White;
        for (var i = 0; i < count; i++)
        {
            var barHeight = Math.Clamp(chart.Values[i] / maximum * plotHeight, 0, plotHeight);
            var bar = new Border { Width = barWidth, Height = barHeight, Background = barBrush };
            Canvas.SetLeft(bar, left + step * i + (step - barWidth) / 2); Canvas.SetTop(bar, top + plotHeight - barHeight); canvas.Children.Add(bar);
            if (i >= chart.Categories.Count) continue;
            var category = new TextBlock { Text = chart.Categories[i], FontSize = 9, Foreground = Brushes.White, Width = Math.Max(38, step * 1.8), TextWrapping = TextWrapping.NoWrap, RenderTransform = new RotateTransform(-45), RenderTransformOrigin = new RelativePoint(0, 0, RelativeUnit.Relative) };
            Canvas.SetLeft(category, left + step * i + step * .15); Canvas.SetTop(category, top + plotHeight + 17); canvas.Children.Add(category);
        }
        if (!string.IsNullOrWhiteSpace(chart.Title))
        {
            var title = new TextBlock { Text = chart.Title, FontSize = 14, FontWeight = FontWeight.Bold, Foreground = Brushes.White, TextAlignment = TextAlignment.Center, Width = plotWidth };
            Canvas.SetLeft(title, left); Canvas.SetTop(title, 0); canvas.Children.Add(title);
        }
        return canvas;
    }

    private void OnPointerWheelChanged(object? sender, PointerWheelEventArgs e)
    {
        if (!e.KeyModifiers.HasFlag(KeyModifiers.Control)) return;
        _zoom = Math.Clamp(_zoom * (e.Delta.Y > 0 ? 1.1 : 1 / 1.1), .5, 2.5);
        _zoomTransform.ScaleX = _zoom;
        _zoomTransform.ScaleY = _zoom;
        _scrollViewer.Offset = new Vector(0, _scrollViewer.Offset.Y);
        UpdateVirtualizedSlides();
        e.Handled = true;
    }

    private void ShowError(Exception error)
    {
        ClearVisuals();
        _slides.Children.Add(new TextBlock { Text = error.Message, Foreground = Brush.Parse("#9F1239"), TextWrapping = TextWrapping.Wrap, Margin = new Thickness(24) });
    }

    private void ClearVisuals()
    {
        foreach (var index in _slideHosts.Keys.ToArray()) ReleaseSlide(index);
        _ownedBitmaps.Clear();
        _slideHosts.Clear();
        _slides.Children.Clear();
        PageCount = 0;
    }

    private static IBrush? ToBrush(PptxFill? fill)
    {
        if (fill is null) return null;
        if (fill.GradientStops.Count == 0) return TryParseBrush(fill.Color);
        var radians = fill.GradientAngle * Math.PI / 180d;
        var dx = Math.Cos(radians) * .5;
        var dy = -Math.Sin(radians) * .5;
        var gradient = new LinearGradientBrush
        {
            StartPoint = new RelativePoint(.5 - dx, .5 - dy, RelativeUnit.Relative),
            EndPoint = new RelativePoint(.5 + dx, .5 + dy, RelativeUnit.Relative)
        };
        foreach (var stop in fill.GradientStops)
            if (Color.TryParse(stop.Color.StartsWith('#') ? stop.Color : "#" + stop.Color, out var color))
                gradient.GradientStops.Add(new GradientStop(color, stop.Position));
        return gradient.GradientStops.Count == 0 ? null : gradient;
    }

    private static IBrush? TryParseBrush(string? color) => color is not null && Color.TryParse(color.StartsWith('#') ? color : "#" + color, out var parsed)
        ? new SolidColorBrush(parsed) : null;

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Clear();
    }

    private sealed class SlideHost(Border slot)
    {
        public Border Slot { get; } = slot;
        public Canvas? Canvas { get; set; }
    }
}
