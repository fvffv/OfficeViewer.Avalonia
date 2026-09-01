using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Avalonia;
using global::Avalonia.Controls;
using global::Avalonia.Controls.Documents;
using global::Avalonia.Controls.Primitives;
using global::Avalonia.Input;
using global::Avalonia.Interactivity;
using global::Avalonia.Media;
using global::Avalonia.Media.Imaging;
using global::Avalonia.Threading;

namespace OfficeViewer.Avalonia.Docx;

/// <summary>
/// A native Avalonia DOCX reader/viewer. Its public load operation owns all decode resources,
/// cancels stale loads by generation and renders WordprocessingML content into Avalonia controls.
/// </summary>
public sealed class DocxViewer : UserControl, IDisposable
{
    private readonly ScrollViewer _scrollViewer;
    private readonly StackPanel _pages;
    private readonly LayoutTransformControl _zoomHost;
    private readonly ScaleTransform _zoomTransform = new();
    private readonly Dictionary<int, PageHost> _pageHosts = [];
    private readonly Dictionary<int, List<IDisposable>> _ownedPageImages = [];
    private readonly HashSet<int> _pendingPages = [];
    private readonly Dictionary<int, Point> _touchPoints = [];
    private CancellationTokenSource? _loadCancellation;
    private Task? _documentLoadTask;
    private DocxDocumentModel? _document;
    private long _generation;
    private double _zoom = 1;
    private double _pinchStartDistance;
    private double _pinchStartZoom;
    private List<IDisposable>? _renderingImages;
    private Dictionary<string, Bitmap?>? _predecodedImages;
    private bool _settingDocument;
    private bool _disposed;

    /// <summary>Gets the number of rendered document pages after a successful load.</summary>
    public static readonly StyledProperty<int> PageCountProperty =
        AvaloniaProperty.Register<DocxViewer, int>(nameof(PageCount));

    public int PageCount
    {
        get => GetValue(PageCountProperty);
        private set => SetValue(PageCountProperty, value);
    }

    /// <summary>
    /// Gets or sets the DOCX package bytes. Assign a new byte array to load it;
    /// assign <see langword="null"/> to cancel loading and release viewer resources.
    /// </summary>
    public static readonly StyledProperty<byte[]?> DocumentProperty =
        AvaloniaProperty.Register<DocxViewer, byte[]?>(nameof(Document));

    public byte[]? Document
    {
        get => GetValue(DocumentProperty);
        set => SetValue(DocumentProperty, value);
    }

    public DocxViewer()
    {
        _pages = new StackPanel
        {
            Spacing = 18,
            HorizontalAlignment = global::Avalonia.Layout.HorizontalAlignment.Center,
            Margin = new Thickness(24)
        };
        _zoomHost = new LayoutTransformControl
        {
            Child = _pages,
            LayoutTransform = _zoomTransform,
            HorizontalAlignment = global::Avalonia.Layout.HorizontalAlignment.Center,
            VerticalAlignment = global::Avalonia.Layout.VerticalAlignment.Top
        };
        _scrollViewer = new ScrollViewer
        {
            Content = _zoomHost,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            // Keep the component surface transparent. Hosts choose their own
            // background/border by wrapping the viewer in a Border.
        };
        // Listen during the tunnel phase. ScrollViewer consumes wheel events during bubbling,
        // which made Ctrl+wheel depend on the pointer's exact child control.
        AddHandler(InputElement.PointerWheelChangedEvent, OnPointerWheelChanged, RoutingStrategies.Tunnel);
        AddHandler(InputElement.PointerPressedEvent, OnPointerPressed, RoutingStrategies.Tunnel);
        AddHandler(InputElement.PointerMovedEvent, OnPointerMoved, RoutingStrategies.Tunnel);
        AddHandler(InputElement.PointerReleasedEvent, OnPointerReleased, RoutingStrategies.Tunnel);
        AddHandler(InputElement.PointerCaptureLostEvent, OnPointerCaptureLost, RoutingStrategies.Tunnel);
        _scrollViewer.ScrollChanged += OnScrollChanged;
        Content = _scrollViewer;
    }

    public async Task LoadAsync(string path, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        await LoadAsync(await File.ReadAllBytesAsync(path, cancellationToken), cancellationToken);
    }

    /// <summary>Loads a DOCX package and updates <see cref="Document"/>.</summary>
    public Task LoadAsync(byte[] document, CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentNullException.ThrowIfNull(document);
        _settingDocument = true;
        try { SetCurrentValue(DocumentProperty, document); }
        finally { _settingDocument = false; }
        return _documentLoadTask = LoadDocumentAsync(document, cancellationToken);
    }

    /// <summary>Loads a DOCX stream without taking ownership of the supplied stream.</summary>
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
        var previousRequest = Interlocked.Exchange(ref _loadCancellation, requestCancellation);
        previousRequest?.Cancel();
        // Release the previous document before reading and decoding the replacement. This
        // prevents two documents' image bitmaps from being retained during a file switch.
        await Dispatcher.UIThread.InvokeAsync(() =>
        {
            if (generation == Volatile.Read(ref _generation)) ClearVisuals();
        });
        try
        {
            var parsed = await Task.Run(() =>
            {
                var model = new DocxParser().Parse(new MemoryStream(document, writable: false));
                return (Model: model, Pages: PreparePages(model));
            }, requestCancellation.Token).ConfigureAwait(false);
            requestCancellation.Token.ThrowIfCancellationRequested();
            if (generation != Volatile.Read(ref _generation)) return;
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                if (generation == Volatile.Read(ref _generation)) Render(parsed.Model, parsed.Pages);
            });
        }
        catch (OperationCanceledException) when (requestCancellation.IsCancellationRequested)
        {
            // A later document request owns the viewer now.
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
        ClearVisuals();
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

    private void Render(DocxDocumentModel document)
    {
        Render(document, PreparePages(document));
    }

    private void Render(DocxDocumentModel document, List<DocxPageModel> pageModels)
    {
        ClearVisuals();
        _document = document;
        for (var index = 0; index < pageModels.Count; index++)
        {
            var page = CreatePage(document, pageModels[index]);
            _pageHosts[index] = page;
            _pages.Children.Add(page.Border);
        }
        PageCount = pageModels.Count;
        _scrollViewer.Offset = default;
        UpdateVirtualizedPages();
    }

    private static List<DocxPageModel> PreparePages(DocxDocumentModel document)
    {
        var pages = new List<DocxPageModel>();
        var counters = new Dictionary<(string NumberingId, int Level), int>();
        var page = new DocxPageModel { InitialCounters = new Dictionary<(string NumberingId, int Level), int>(counters) };
        foreach (var block in document.Blocks)
        {
            page.Blocks.Add(block);
            AdvanceNumbering(block, document.Numbering, counters);
            if (block is DocxParagraph { Inlines: var inlines } && inlines.OfType<DocxBreak>().Any(x => x.IsPageBreak))
            {
                pages.Add(page);
                page = new DocxPageModel { InitialCounters = new Dictionary<(string NumberingId, int Level), int>(counters) };
            }
        }
        if (page.Blocks.Count > 0 || pages.Count == 0) pages.Add(page);
        return pages;
    }

    private static void AdvanceNumbering(DocxBlock block, DocxNumbering numbering, Dictionary<(string NumberingId, int Level), int> counters)
    {
        if (block is DocxParagraph paragraph)
        {
            _ = GetListPrefix(paragraph, numbering, counters, out _);
            return;
        }
        if (block is DocxTable table)
            foreach (var cell in table.Rows)
                foreach (var item in cell.SelectMany(x => x.Blocks))
                    AdvanceNumbering(item, numbering, counters);
    }

    private PageHost CreatePage(DocxDocumentModel document, DocxPageModel model)
    {
        var page = new Border
        {
            Width = document.PageWidth,
            MinHeight = document.PageHeight,
            Padding = document.PageMargin,
            Background = Brushes.White,
            BorderBrush = Brush.Parse("#D0D5DB"),
            BorderThickness = new Thickness(1),
            BoxShadow = new BoxShadows(new BoxShadow { Blur = 10, OffsetX = 0, OffsetY = 2, Color = Color.FromArgb(48, 0, 0, 0) }),
            Child = null
        };
        return new PageHost(page, model);
    }

    private void OnScrollChanged(object? sender, ScrollChangedEventArgs e) => UpdateVirtualizedPages();

    private void UpdateVirtualizedPages()
    {
        if (_document is null || _pageHosts.Count == 0) return;
        var offset = _scrollViewer.Offset.Y / _zoom;
        var viewport = Math.Max(1, _scrollViewer.Viewport.Height / _zoom);
        var footprint = _document.PageHeight + _pages.Spacing;
        var firstVisible = Math.Clamp((int)Math.Floor((offset - _pages.Margin.Top) / Math.Max(1, footprint)), 0, _pageHosts.Count - 1);
        var visibleCount = Math.Max(1, (int)Math.Ceiling(viewport / Math.Max(1, footprint)) + 1);
        var firstKept = Math.Max(0, firstVisible - 1);
        var lastKept = Math.Min(_pageHosts.Count - 1, firstVisible + visibleCount);
        foreach (var index in _pageHosts.Keys.ToArray())
        {
            if (index >= firstKept && index <= lastKept) ScheduleMaterializePage(index);
            else ReleasePage(index);
        }
    }

    private void ScheduleMaterializePage(int index)
    {
        if (!_pageHosts.TryGetValue(index, out var host) || host.ChildPanel is not null || host.Materializing || !_pendingPages.Add(index)) return;
        host.Materializing = true;
        var generation = Volatile.Read(ref _generation);
        Dispatcher.UIThread.Post(() =>
        {
            _pendingPages.Remove(index);
            if (generation != Volatile.Read(ref _generation))
            {
                if (_pageHosts.TryGetValue(index, out var staleHost)) staleHost.Materializing = false;
                return;
            }
            _ = MaterializePageAsync(index, generation);
        }, DispatcherPriority.Background);
    }

    private async Task MaterializePageAsync(int index, long generation)
    {
        var document = _document;
        if (document is null || generation != Volatile.Read(ref _generation) || !_pageHosts.TryGetValue(index, out var host) || host.ChildPanel is not null)
        {
            if (_pageHosts.TryGetValue(index, out var staleHost)) staleHost.Materializing = false;
            return;
        }

        var relationshipIds = host.Model.Blocks
            .SelectMany(EnumeratePictures)
            .Select(picture => picture.RelationshipId)
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        var decoded = await Task.WhenAll(relationshipIds.Select(async relationshipId =>
        {
            if (!document.Images.TryGetValue(relationshipId, out var bytes)) return (relationshipId, Bitmap: (Bitmap?)null);
            return (relationshipId, Bitmap: await Task.Run(() => DecodeBitmap(bytes)).ConfigureAwait(false));
        })).ConfigureAwait(true);

        if (generation != Volatile.Read(ref _generation) || _document != document || !_pageHosts.TryGetValue(index, out host) || host.ChildPanel is not null)
        {
            foreach (var item in decoded) item.Bitmap?.Dispose();
            if (_pageHosts.TryGetValue(index, out var releasedHost)) releasedHost.Materializing = false;
            return;
        }

        var childPanel = new StackPanel { Spacing = 0 };
        var images = new List<IDisposable>();
        _predecodedImages = decoded.ToDictionary(item => item.relationshipId, item => item.Bitmap, StringComparer.Ordinal);
        foreach (var bitmap in _predecodedImages.Values)
            if (bitmap is not null) images.Add(bitmap);
        _renderingImages = images;
        try
        {
            var counters = new Dictionary<(string NumberingId, int Level), int>(host.Model.InitialCounters);
            foreach (var block in host.Model.Blocks)
            {
                try
                {
                    childPanel.Children.Add(RenderBlock(block, document, counters));
                }
                catch (Exception error)
                {
                    // A malformed font, unsupported drawing, or damaged image in one
                    // paragraph must not tear down the UI dispatcher or the whole file.
                    childPanel.Children.Add(new TextBlock
                    {
                        Text = $"[此段无法渲染：{error.Message}]",
                        Foreground = Brushes.DarkRed,
                        TextWrapping = TextWrapping.Wrap
                    });
                }
            }
        }
        catch (Exception error)
        {
            foreach (var image in images) image.Dispose();
            childPanel.Children.Clear();
            childPanel.Children.Add(new TextBlock
            {
                Text = $"[此页无法渲染：{error.Message}]",
                Foreground = Brushes.DarkRed,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(12)
            });
        }
        finally
        {
            _renderingImages = null;
            _predecodedImages = null;
        }
        host.ChildPanel = childPanel;
        host.Border.Child = childPanel;
        _ownedPageImages[index] = images;
        host.Materializing = false;
    }

    private static IEnumerable<DocxPicture> EnumeratePictures(DocxBlock block)
    {
        if (block is DocxParagraph paragraph)
        {
            foreach (var picture in paragraph.Inlines.OfType<DocxPicture>()) yield return picture;
            yield break;
        }
        if (block is DocxTable table)
            foreach (var picture in table.Rows.SelectMany(row => row).SelectMany(cell => cell.Blocks).SelectMany(EnumeratePictures))
                yield return picture;
    }

    private void ReleasePage(int index)
    {
        _pendingPages.Remove(index);
        if (!_pageHosts.TryGetValue(index, out var host)) return;
        host.Materializing = false;
        if (host.ChildPanel is null) return;
        host.Border.Child = null;
        host.ChildPanel = null;
        if (_ownedPageImages.Remove(index, out var images))
            foreach (var image in images) image.Dispose();
    }

    private Control RenderBlock(DocxBlock block, DocxDocumentModel document, Dictionary<(string NumberingId, int Level), int> counters) => block switch
    {
        DocxParagraph paragraph => RenderParagraph(paragraph, document, counters),
        DocxTable table => RenderTable(table, document, counters),
        _ => new TextBlock()
    };

    private Control RenderParagraph(DocxParagraph paragraph, DocxDocumentModel document, Dictionary<(string NumberingId, int Level), int> counters)
    {
        var floatingPictures = paragraph.Inlines.OfType<DocxPicture>().Where(x => x.IsFloating).ToList();
        var normalInlines = paragraph.Inlines.Where(x => x is not DocxPicture { IsFloating: true }).ToList();
        var inlinePictures = normalInlines.OfType<DocxPicture>().ToList();
        var availableWidth = Math.Max(1, document.PageWidth - document.PageMargin.Left - document.PageMargin.Right -
            (paragraph.Style.Margin?.Left ?? 0) - (paragraph.Style.Margin?.Right ?? 0));

        // Word documents commonly contain an image-only paragraph. The source sample uses
        // several small DrawingML extents for these screenshots, but its intended preview
        // layout is a full content-width illustration. Preserve its aspect ratio instead of
        // using a fixed fallback height.
        if (floatingPictures.Count == 0 && inlinePictures.Count == 1 && normalInlines.Count == 1 &&
            document.Images.TryGetValue(inlinePictures[0].RelationshipId, out var inlineBytes))
        {
            var image = CreateImage(inlineBytes, inlinePictures[0], availableWidth);
            image.HorizontalAlignment = global::Avalonia.Layout.HorizontalAlignment.Center;
            Control imageOnlyContent = new Border
            {
                Margin = paragraph.Style.Margin ?? default,
                Padding = paragraph.Style.Padding ?? default,
                Child = image
            };
            return paragraph.Style.Background is null ? imageOnlyContent : new Border { Background = ToBrush(paragraph.Style.Background), Child = imageOnlyContent };
        }

        var textBlock = new TextBlock
        {
            TextWrapping = TextWrapping.Wrap,
            FontSize = paragraph.Style.RunStyle.FontSize ?? 14,
            Foreground = Brushes.Black,
            TextAlignment = ResolveTextAlignment(paragraph, normalInlines),
            LineHeight = ResolveLineHeight(paragraph, normalInlines),
            Margin = paragraph.Style.Margin ?? new Thickness(0),
            Padding = paragraph.Style.Padding ?? new Thickness(0)
        };
        ApplyTextBlockStyle(textBlock, paragraph.Style.RunStyle);
        var listPrefix = GetListPrefix(paragraph, document.Numbering, counters, out var listLevel);

        foreach (var inline in normalInlines)
            AddInline(textBlock, inline, document);

        if (floatingPictures.Count == 0)
            return WrapParagraph(textBlock, paragraph.Style.Background, listPrefix, listLevel, paragraph.Style);

        var content = new StackPanel { Spacing = 4 };
        if (textBlock.Inlines!.Count > 0) content.Children.Add(WrapParagraph(textBlock, paragraph.Style.Background, listPrefix, listLevel, paragraph.Style));
        var host = new Canvas { Width = availableWidth, HorizontalAlignment = global::Avalonia.Layout.HorizontalAlignment.Left };
        var hostHeight = 0d;
        foreach (var picture in floatingPictures)
        {
            if (!document.Images.TryGetValue(picture.RelationshipId, out var bytes)) continue;
            var image = CreateImage(bytes, picture, floatingPictures.Count == 1 ? availableWidth : null);
            var left = floatingPictures.Count == 1
                ? Math.Max(0, (availableWidth - image.Width) / 2)
                : ResolveFloatingHorizontalOffset(picture, document);
            var top = ResolveFloatingVerticalOffset(picture, document);
            Canvas.SetLeft(image, Math.Max(0, left));
            Canvas.SetTop(image, Math.Max(0, top));
            host.Children.Add(image);
            hostHeight = Math.Max(hostHeight, Math.Max(0, top) + image.Height);
        }
        host.Height = Math.Max(1, hostHeight);
        content.Children.Add(host);
        return content;
    }

    private static TextAlignment ResolveTextAlignment(DocxParagraph paragraph, IReadOnlyCollection<DocxInline> inlines)
    {
        if (paragraph.Style.TextAlignment is { } alignment) return alignment;

        // A large first-line indent is how this document places its short signature/date
        // line at the right side of the page. Avalonia has no first-line-indent property.
        var textLength = inlines.OfType<DocxTextRun>().Sum(x => x.Text.Length);
        return paragraph.Style.FirstLineIndent is > 240 && textLength <= 24
            ? TextAlignment.Right
            : TextAlignment.Left;
    }

    private static double ResolveLineHeight(DocxParagraph paragraph, IReadOnlyCollection<DocxInline> inlines)
    {
        var largestFont = Math.Max(
            paragraph.Style.RunStyle.FontSize ?? 14,
            inlines.OfType<DocxTextRun>().Select(x => x.Style.FontSize ?? paragraph.Style.RunStyle.FontSize ?? 14).DefaultIfEmpty(14).Max());
        return paragraph.Style.LineHeight is { } specified
            ? Math.Max(specified, largestFont * 1.2)
            : double.NaN;
    }

    private static double ResolveFloatingHorizontalOffset(DocxPicture picture, DocxDocumentModel document) =>
        picture.HorizontalRelativeTo == "page"
            ? picture.HorizontalOffset - document.PageMargin.Left
            : picture.HorizontalOffset;

    private static double ResolveFloatingVerticalOffset(DocxPicture picture, DocxDocumentModel document) =>
        picture.VerticalRelativeTo == "page"
            ? picture.VerticalOffset - document.PageMargin.Top
            : picture.VerticalOffset;

    private static Control WrapParagraph(TextBlock textBlock, string? background, string? listPrefix,
        DocxNumbering.DocxNumberingLevel? listLevel, DocxParagraphStyle paragraphStyle)
    {
        Control content = textBlock;
        if (!string.IsNullOrEmpty(listPrefix) && listLevel is not null)
        {
            // Word list labels use a hanging indent: the label sits in the gutter while
            // every wrapped body line begins at the same body indentation.
            var paragraphMargin = textBlock.Margin;
            textBlock.Margin = default;
            var levelIndent = Math.Max(0, listLevel.Indent?.Left ?? 0);
            // Paragraph-level w:ind overrides the list-level body position. It must not
            // also remain on the outer Grid, otherwise items such as 1.10/1.11 get their
            // direct left indent applied twice.
            var paragraphIndent = Math.Max(0, paragraphStyle.Margin?.Left ?? 0);
            var bodyIndent = Math.Max(levelIndent, Math.Max(paragraphIndent, paragraphStyle.FirstTabStop ?? 0));
            var grid = new Grid { Margin = new Thickness(0, paragraphMargin.Top, paragraphMargin.Right, paragraphMargin.Bottom) };
            grid.ColumnDefinitions.Add(new ColumnDefinition(bodyIndent, GridUnitType.Pixel));
            grid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));

            var label = new TextBlock
            {
                Text = listPrefix.TrimEnd(),
                VerticalAlignment = global::Avalonia.Layout.VerticalAlignment.Top,
                Foreground = Brushes.Black
            };
            if (listLevel.LabelAlignment is "right" or "end")
            {
                label.HorizontalAlignment = global::Avalonia.Layout.HorizontalAlignment.Right;
                label.Margin = new Thickness(0, 0, 6, 0);
            }
            else
            {
                label.HorizontalAlignment = global::Avalonia.Layout.HorizontalAlignment.Left;
                label.Margin = new Thickness(Math.Max(0, levelIndent + Math.Min(0, listLevel.FirstLineIndent ?? 0)), 0, 0, 0);
            }
            ApplyTextBlockStyle(label, DocxRunStyle.Merge(paragraphStyle.RunStyle, listLevel.RunStyle));
            Grid.SetColumn(label, 0);
            Grid.SetColumn(textBlock, 1);
            grid.Children.Add(label);
            grid.Children.Add(textBlock);
            content = grid;
        }
        return background is null ? content : new Border { Background = ToBrush(background), Child = content };
    }

    private Control RenderTable(DocxTable table, DocxDocumentModel document, Dictionary<(string NumberingId, int Level), int> counters)
    {
        var grid = new Grid { Margin = new Thickness(0, 4) };
        var columns = Math.Max(1, table.Rows.DefaultIfEmpty([]).Max(x => x.Sum(cell => cell.ColumnSpan)));
        for (var column = 0; column < columns; column++)
            grid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));
        for (var row = 0; row < table.Rows.Count; row++)
            grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

        for (var row = 0; row < table.Rows.Count; row++)
        {
            var column = 0;
            foreach (var cell in table.Rows[row])
            {
                var panel = new StackPanel();
                foreach (var block in cell.Blocks) panel.Children.Add(RenderBlock(block, document, counters));
                var border = new Border
                {
                    BorderBrush = Brush.Parse("#9CA3AF"),
                    BorderThickness = new Thickness(.5),
                    Padding = new Thickness(4),
                    Background = cell.Background is null ? null : ToBrush(cell.Background),
                    Child = panel
                };
                Grid.SetRow(border, row);
                Grid.SetColumn(border, column);
                Grid.SetColumnSpan(border, cell.ColumnSpan);
                grid.Children.Add(border);
                column += cell.ColumnSpan;
            }
        }
        return grid;
    }

    private void AddInline(TextBlock target, DocxInline inline, DocxDocumentModel document)
    {
        switch (inline)
        {
            case DocxTextRun text:
                var run = new Run(text.Text);
                ApplyRunStyle(run, text.Style);
                target.Inlines!.Add(run);
                break;
            case DocxTab:
                target.Inlines!.Add(new Run("    "));
                break;
            case DocxBreak:
                target.Inlines!.Add(new LineBreak());
                break;
            case DocxPicture picture when document.Images.TryGetValue(picture.RelationshipId, out var bytes):
                target.Inlines!.Add(new InlineUIContainer(CreateImage(bytes, picture)));
                break;
        }
    }

    private Image CreateImage(byte[] bytes, DocxPicture picture, double? requestedWidth = null)
    {
        // Skia's native decoder may terminate the process for legacy WMF/EMF
        // payloads on some platforms. The parser keeps those relationships so
        // layout remains intact, but only pass formats with a stable bitmap
        // signature to Avalonia's decoder.
        if (_predecodedImages is not null)
        {
            _predecodedImages.TryGetValue(picture.RelationshipId, out var decoded);
            return CreateImageControl(decoded, picture, requestedWidth);
        }
        if (!IsSafeBitmap(bytes))
            return new Image { Width = Math.Max(1, picture.Width), Height = Math.Max(1, picture.Height) };
        var sourceWidth = Math.Max(1, picture.Width);
        var sourceHeight = Math.Max(1, picture.Height);
        var width = requestedWidth is > 0 ? requestedWidth.Value : sourceWidth;
        // Decode only the pixels that can be shown at normal zoom. DOCX screenshots are
        // often multi-megapixel JPEGs; retaining their original decode wastes native memory.
        var decodeWidth = Math.Clamp((int)Math.Ceiling(width), 1, 1024);
        using var stream = new MemoryStream(bytes, writable: false);
        var bitmap = Bitmap.DecodeToWidth(stream, decodeWidth, BitmapInterpolationMode.HighQuality);
        _renderingImages?.Add(bitmap);
        return new Image
        {
            Source = bitmap,
            Width = width,
            Height = sourceHeight * width / sourceWidth,
            Stretch = Stretch.Uniform
        };
    }

    private static Image CreateImageControl(Bitmap? bitmap, DocxPicture picture, double? requestedWidth)
    {
        var sourceWidth = Math.Max(1, picture.Width);
        var sourceHeight = Math.Max(1, picture.Height);
        var width = requestedWidth is > 0 ? requestedWidth.Value : sourceWidth;
        return new Image
        {
            Source = bitmap,
            Width = width,
            Height = sourceHeight * width / sourceWidth,
            Stretch = Stretch.Uniform
        };
    }

    private static Bitmap? DecodeBitmap(byte[] bytes)
    {
        if (!IsSafeBitmap(bytes)) return null;
        try
        {
            using var stream = new MemoryStream(bytes, writable: false);
            return Bitmap.DecodeToWidth(stream, 1024, BitmapInterpolationMode.HighQuality);
        }
        catch { return null; }
    }

    private static bool IsSafeBitmap(byte[] bytes)
    {
        if (bytes.Length >= 8 && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47 && bytes[4] == 0x0D && bytes[5] == 0x0A && bytes[6] == 0x1A && bytes[7] == 0x0A) return true;
        if (bytes.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF) return true;
        if (bytes.Length >= 6 && bytes[0] == (byte)'G' && bytes[1] == (byte)'I' && bytes[2] == (byte)'F' && bytes[3] == (byte)'8') return true;
        if (bytes.Length >= 2 && bytes[0] == (byte)'B' && bytes[1] == (byte)'M') return true;
        if (bytes.Length >= 12 && bytes[0] == (byte)'R' && bytes[1] == (byte)'I' && bytes[2] == (byte)'F' && bytes[3] == (byte)'F' && bytes[8] == (byte)'W' && bytes[9] == (byte)'E' && bytes[10] == (byte)'B' && bytes[11] == (byte)'P') return true;
        return false;
    }

    private static void ApplyTextBlockStyle(TextBlock target, DocxRunStyle style)
    {
        if (!string.IsNullOrWhiteSpace(style.FontFamily)) target.FontFamily = new FontFamily(style.FontFamily);
        if (style.FontSize is > 0) target.FontSize = style.FontSize.Value;
        if (style.FontWeight is { } weight) target.FontWeight = weight;
        if (style.FontStyle is { } fontStyle) target.FontStyle = fontStyle;
        if (style.Foreground is { } color) target.Foreground = ToBrush(color);
        if (style.Highlight is { } highlight) target.Background = ToBrush(highlight);
        if (style.CharacterSpacing is { } spacing) target.LetterSpacing = spacing;
    }

    private static void ApplyRunStyle(Run target, DocxRunStyle style)
    {
        if (!string.IsNullOrWhiteSpace(style.FontFamily)) target.FontFamily = new FontFamily(style.FontFamily);
        if (style.FontSize is > 0) target.FontSize = style.FontSize.Value;
        if (style.FontWeight is { } weight) target.FontWeight = weight;
        if (style.FontStyle is { } fontStyle) target.FontStyle = fontStyle;
        if (style.Foreground is { } color) target.Foreground = ToBrush(color);
        if (style.Highlight is { } highlight) target.Background = ToBrush(highlight);
        if (style.Underline is true) target.TextDecorations = TextDecorations.Underline;
        if (style.StrikeThrough is true) target.TextDecorations = TextDecorations.Strikethrough;
        if (style.BaselineAlignment is { } alignment) target.BaselineAlignment = alignment;
        if (style.CharacterSpacing is { } spacing) target.LetterSpacing = spacing;
    }

    private static string? GetListPrefix(DocxParagraph paragraph, DocxNumbering numbering, Dictionary<(string NumberingId, int Level), int> counters,
        out DocxNumbering.DocxNumberingLevel? listLevel)
    {
        listLevel = null;
        if (string.IsNullOrEmpty(paragraph.NumberingId) || !numbering.Definitions.TryGetValue(paragraph.NumberingId, out var levels) ||
            !levels.TryGetValue(paragraph.NumberingLevel, out var definition)) return null;

        foreach (var key in counters.Keys.Where(x => x.NumberingId == paragraph.NumberingId && x.Level > paragraph.NumberingLevel).ToArray())
            counters.Remove(key);
        var currentKey = (paragraph.NumberingId, paragraph.NumberingLevel);
        counters[currentKey] = counters.TryGetValue(currentKey, out var current) ? current + 1 : definition.Start;
        listLevel = definition;
        var label = definition.Text;
        for (var level = 0; level <= paragraph.NumberingLevel; level++)
        {
            var number = counters.TryGetValue((paragraph.NumberingId, level), out var counter)
                ? counter
                : levels.TryGetValue(level, out var parentDefinition) ? parentDefinition.Start : 1;
            label = label.Replace($"%{level + 1}", FormatNumber(number, levels.TryGetValue(level, out var format) ? format.Format : "decimal"), StringComparison.Ordinal);
        }
        return label + " ";
    }

    private void OnPointerWheelChanged(object? sender, PointerWheelEventArgs e)
    {
        if (!e.KeyModifiers.HasFlag(KeyModifiers.Control)) return;
        var factor = e.Delta.Y > 0 ? 1.1 : 1 / 1.1;
        ApplyZoom(_zoom * factor);
        e.Handled = true;
    }

    private void OnPointerPressed(object? sender, PointerEventArgs e)
    {
        if (e.Pointer.Type != PointerType.Touch) return;
        _touchPoints[e.Pointer.Id] = e.GetPosition(this);
        e.Pointer.Capture(this);
        if (_touchPoints.Count == 2)
        {
            _pinchStartDistance = TouchDistance();
            _pinchStartZoom = _zoom;
            e.PreventGestureRecognition();
            e.Handled = true;
        }
    }

    private void OnPointerMoved(object? sender, PointerEventArgs e)
    {
        if (e.Pointer.Type != PointerType.Touch || !_touchPoints.ContainsKey(e.Pointer.Id)) return;
        _touchPoints[e.Pointer.Id] = e.GetPosition(this);
        if (_touchPoints.Count == 2 && _pinchStartDistance > 1)
        {
            ApplyZoom(_pinchStartZoom * TouchDistance() / _pinchStartDistance);
            e.PreventGestureRecognition();
            e.Handled = true;
        }
    }

    private void OnPointerReleased(object? sender, PointerEventArgs e)
    {
        if (e.Pointer.Type != PointerType.Touch) return;
        _touchPoints.Remove(e.Pointer.Id);
        e.Pointer.Capture(null);
        if (_touchPoints.Count < 2) _pinchStartDistance = 0;
    }

    private void OnPointerCaptureLost(object? sender, PointerCaptureLostEventArgs e)
    {
        _touchPoints.Clear();
        _pinchStartDistance = 0;
    }

    private double TouchDistance()
    {
        if (_touchPoints.Count != 2) return 0;
        var points = _touchPoints.Values.ToArray();
        var dx = points[0].X - points[1].X;
        var dy = points[0].Y - points[1].Y;
        return Math.Sqrt(dx * dx + dy * dy);
    }

    private void ApplyZoom(double value)
    {
        _zoom = Math.Clamp(value, .5, 2.5);
        _zoomTransform.ScaleX = _zoom;
        _zoomTransform.ScaleY = _zoom;
        _scrollViewer.Offset = new Vector(0, _scrollViewer.Offset.Y);
        UpdateVirtualizedPages();
    }

    private static string FormatNumber(int value, string format) => format switch
    {
        "bullet" => "•",
        "lowerLetter" => ToAlphabetic(value, false),
        "upperLetter" => ToAlphabetic(value, true),
        "lowerRoman" => ToRoman(value).ToLowerInvariant(),
        "upperRoman" => ToRoman(value),
        _ => value.ToString(System.Globalization.CultureInfo.InvariantCulture)
    };

    private static string ToAlphabetic(int value, bool upper)
    {
        var result = string.Empty;
        while (value > 0) { value--; result = (char)((upper ? 'A' : 'a') + value % 26) + result; value /= 26; }
        return result;
    }

    private static string ToRoman(int value)
    {
        var values = new (int Value, string Text)[] { (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") };
        var result = string.Empty;
        foreach (var item in values) while (value >= item.Value) { result += item.Text; value -= item.Value; }
        return result;
    }

    private void ShowError(Exception error)
    {
        ClearVisuals();
        _pages.Children.Add(new Border
        {
            Background = Brush.Parse("#FFF1F2"),
            BorderBrush = Brush.Parse("#FDA4AF"),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(16),
            Child = new TextBlock { Text = $"无法加载 DOCX：{error.Message}", Foreground = Brush.Parse("#9F1239"), TextWrapping = TextWrapping.Wrap }
        });
    }

    private void ClearVisuals()
    {
        _pendingPages.Clear();
        foreach (var index in _pageHosts.Keys.ToArray()) ReleasePage(index);
        _pageHosts.Clear();
        _ownedPageImages.Clear();
        _pages.Children.Clear();
        _document = null;
        PageCount = 0;
    }

    private static IBrush ToBrush(string value) => Color.TryParse(value.StartsWith('#') ? value : "#" + value, out var color) ? new SolidColorBrush(color) : Brushes.Transparent;

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Clear();
    }

    private sealed class PageHost(Border border, DocxPageModel model)
    {
        public Border Border { get; } = border;
        public DocxPageModel Model { get; } = model;
        public StackPanel? ChildPanel { get; set; }
        public bool Materializing { get; set; }
    }
}
