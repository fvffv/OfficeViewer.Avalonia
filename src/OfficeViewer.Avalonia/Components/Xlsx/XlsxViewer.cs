using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using Avalonia;
using global::Avalonia.Controls;
using global::Avalonia.Controls.Documents;
using global::Avalonia.Controls.Primitives;
using global::Avalonia.Data;
using global::Avalonia.Input;
using global::Avalonia.Interactivity;
using global::Avalonia.Layout;
using global::Avalonia.Media;
using global::Avalonia.Media.Imaging;
using global::Avalonia.Threading;

namespace OfficeViewer.Avalonia.Xlsx;

/// <summary>Read-only, AOT-safe XLSX sheet viewer with one active materialized sheet.</summary>
public sealed class XlsxViewer : UserControl, IDisposable
{
    private readonly ScrollViewer _scrollViewer;
    private readonly Border _sheetHost = new();
    private readonly List<IDisposable> _ownedImages = [];
    private XlsxWorkbookModel? _workbook;
    private CancellationTokenSource? _loadCancellation;
    private Task? _documentLoadTask;
    private long _generation;
    private int _activeSheetIndex = -1;
    private bool _settingDocument;
    private bool _disposed;

    /// <summary>Gets the number of worksheets in the successfully loaded workbook.</summary>
    public static readonly StyledProperty<int> SheetCountProperty =
        AvaloniaProperty.Register<XlsxViewer, int>(nameof(SheetCount));

    /// <summary>
    /// Gets or sets the zero-based worksheet to display. Bind this property from the host;
    /// no worksheet selector is rendered by the viewer itself.
    /// </summary>
    public static readonly StyledProperty<int> SelectedSheetIndexProperty =
        AvaloniaProperty.Register<XlsxViewer, int>(nameof(SelectedSheetIndex), defaultBindingMode: BindingMode.TwoWay);

    public int SheetCount
    {
        get => GetValue(SheetCountProperty);
        private set => SetValue(SheetCountProperty, value);
    }

    public int SelectedSheetIndex
    {
        get => GetValue(SelectedSheetIndexProperty);
        set => SetValue(SelectedSheetIndexProperty, value);
    }

    /// <summary>
    /// Gets or sets the XLSX package bytes. Assign a new byte array to load it;
    /// assign <see langword="null"/> to cancel loading and release viewer resources.
    /// </summary>
    public static readonly StyledProperty<byte[]?> DocumentProperty =
        AvaloniaProperty.Register<XlsxViewer, byte[]?>(nameof(Document));

    public byte[]? Document
    {
        get => GetValue(DocumentProperty);
        set => SetValue(DocumentProperty, value);
    }

    public XlsxViewer()
    {
        _scrollViewer = new ScrollViewer
        {
            Content = _sheetHost,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        };
        Content = _scrollViewer; // no viewer background: hosts own the surrounding surface
        AddHandler(InputElement.PointerWheelChangedEvent, OnPointerWheelChanged, RoutingStrategies.Tunnel);
    }

    public async Task LoadAsync(string path, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        await LoadAsync(await File.ReadAllBytesAsync(path, cancellationToken), cancellationToken);
    }

    /// <summary>Loads an XLSX package and updates <see cref="Document"/>.</summary>
    public Task LoadAsync(byte[] document, CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentNullException.ThrowIfNull(document);
        _settingDocument = true;
        try { SetCurrentValue(DocumentProperty, document); }
        finally { _settingDocument = false; }
        return _documentLoadTask = LoadDocumentAsync(document, cancellationToken);
    }

    /// <summary>Loads an XLSX stream without taking ownership of the supplied stream.</summary>
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
        await Dispatcher.UIThread.InvokeAsync(() =>
        {
            _workbook = null;
            ClearVisuals();
        });
        try
        {
            var workbook = await Task.Run(() => new XlsxParser().Parse(document), requestCancellation.Token).ConfigureAwait(false);
            requestCancellation.Token.ThrowIfCancellationRequested();
            if (generation != Volatile.Read(ref _generation)) return;
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                if (generation != Volatile.Read(ref _generation)) return;
                _workbook = workbook;
                SheetCount = workbook.Sheets.Count;
                var displayIndex = workbook.Sheets.Count == 0
                    ? -1
                    : Math.Clamp(SelectedSheetIndex, 0, workbook.Sheets.Count - 1);
                if (displayIndex < 0) return;
                if (displayIndex != SelectedSheetIndex)
                    SetCurrentValue(SelectedSheetIndexProperty, displayIndex);
                else
                    ShowSheet(displayIndex);
            });
        }
        catch (OperationCanceledException) when (requestCancellation.IsCancellationRequested) { }
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
        _workbook = null;
        ClearVisuals();
    }

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == DocumentProperty && !_settingDocument)
        {
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
        else if (change.Property == SelectedSheetIndexProperty)
            ShowSheet(change.GetNewValue<int>());
    }

    private void ShowSheet(int index)
    {
        if (_workbook is null || index < 0 || index >= _workbook.Sheets.Count) return;
        if (_activeSheetIndex == index) return;
        _activeSheetIndex = index;
        // Replacing the Grid disposes the whole control tree for the prior sheet. The
        // model is compact strings/styles only, so only one native visual sheet exists.
        ReleaseImages();
        _sheetHost.Child = BuildSheet(_workbook.Sheets[index]);
        _scrollViewer.Offset = default;
    }

    private Grid BuildSheet(XlsxSheetModel sheet)
    {
        var grid = new Grid { Background = Brushes.White, Margin = new Thickness(8) };
        grid.RowDefinitions.Add(new RowDefinition(new GridLength(24)));
        grid.ColumnDefinitions.Add(new ColumnDefinition(new GridLength(52)));
        for (var column = 0; column < sheet.ColumnCount; column++)
            grid.ColumnDefinitions.Add(new ColumnDefinition(new GridLength(sheet.ColumnWidths.TryGetValue(column, out var width) ? width : sheet.DefaultColumnWidth)));
        for (var row = 0; row < sheet.RowCount; row++)
            grid.RowDefinitions.Add(new RowDefinition(new GridLength(sheet.RowHeights.TryGetValue(row, out var height) ? height : sheet.DefaultRowHeight)));

        for (var column = 0; column < sheet.ColumnCount; column++)
        {
            var header = Header(ColumnName(column));
            Grid.SetRow(header, 0); Grid.SetColumn(header, column + 1); grid.Children.Add(header);
        }
        for (var row = 0; row < sheet.RowCount; row++)
        {
            var header = Header((row + 1).ToString());
            Grid.SetRow(header, row + 1); Grid.SetColumn(header, 0); grid.Children.Add(header);
        }
        grid.Children.Add(Header(string.Empty));

        for (var row = 0; row < sheet.RowCount; row++)
            for (var column = 0; column < sheet.ColumnCount; column++)
            {
                if (sheet.Cells.TryGetValue((row, column), out var cell) && cell.IsMergedChild) continue;
                var content = sheet.Cells.TryGetValue((row, column), out cell) ? CreateCell(cell) : EmptyCell();
                Grid.SetRow(content, row + 1); Grid.SetColumn(content, column + 1);
                if (cell is { RowSpan: > 1 }) Grid.SetRowSpan(content, cell.RowSpan);
                if (cell is { ColumnSpan: > 1 }) Grid.SetColumnSpan(content, cell.ColumnSpan);
                grid.Children.Add(content);
            }
        if (sheet.Images.Count > 0) AddImageOverlay(grid, sheet);
        return grid;
    }

    private static Border Header(string text) => new()
    {
        Background = Brush.Parse("#F3F4F6"), BorderBrush = Brush.Parse("#D1D5DB"), BorderThickness = new Thickness(.5),
        Child = new TextBlock { Text = text, FontSize = 12, Foreground = Brushes.Black, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center }
    };

    private static Border EmptyCell() => new() { BorderBrush = Brush.Parse("#D1D5DB"), BorderThickness = new Thickness(.5) };

    private static Border CreateCell(XlsxCellModel cell)
    {
        var style = cell.Style;
        var text = new TextBlock
        {
            TextWrapping = style.Wrap ? TextWrapping.Wrap : TextWrapping.NoWrap,
            TextAlignment = style.HorizontalAlignment,
            VerticalAlignment = style.VerticalAlignment,
            Padding = new Thickness(4, 2),
            Foreground = ToBrush(style.Foreground) ?? Brushes.Black
        };
        if (!string.IsNullOrWhiteSpace(style.FontFamily)) text.FontFamily = new FontFamily(style.FontFamily);
        if (style.FontSize is > 0) text.FontSize = style.FontSize.Value;
        if (style.FontWeight is { } weight) text.FontWeight = weight;
        if (style.FontStyle is { } fontStyle) text.FontStyle = fontStyle;
        if (cell.Runs.Count == 0) text.Text = cell.Text;
        else foreach (var item in cell.Runs)
        {
            var run = new Run(item.Text) { Foreground = ToBrush(item.Foreground) ?? ToBrush(style.Foreground) ?? Brushes.Black };
            if (!string.IsNullOrWhiteSpace(item.FontFamily)) run.FontFamily = new FontFamily(item.FontFamily);
            if (item.FontSize is > 0) run.FontSize = item.FontSize.Value;
            if (item.FontWeight is { } runWeight) run.FontWeight = runWeight;
            if (item.FontStyle is { } runFontStyle) run.FontStyle = runFontStyle;
            if (item.Underline) run.TextDecorations = TextDecorations.Underline;
            text.Inlines!.Add(run);
        }
        // In an Excel grid, text can visually spill into an adjacent blank cell. Native
        // Grid children cannot do that reliably, especially for the final used column.
        // Scale only down for every non-wrapping value so labels and numeric totals are
        // fully readable without changing the document's row heights or overlapping rows.
        Control content = !style.Wrap && !string.IsNullOrEmpty(cell.Text)
            ? new Viewbox { Child = text, Stretch = Stretch.Uniform, StretchDirection = StretchDirection.DownOnly }
            : text;
        return new Border
        {
            Background = ToBrush(style.Background),
            BorderBrush = ToBrush(style.BorderColor) ?? Brush.Parse("#D1D5DB"),
            BorderThickness = style.BorderThickness == default ? new Thickness(.5) : style.BorderThickness,
            Child = content
        };
    }

    private void AddImageOverlay(Grid grid, XlsxSheetModel sheet)
    {
        if (_workbook is null) return;
        var bodyWidth = 0d;
        for (var column = 0; column < sheet.ColumnCount; column++)
            bodyWidth += sheet.ColumnWidths.TryGetValue(column, out var width) ? width : sheet.DefaultColumnWidth;
        var bodyHeight = 0d;
        for (var row = 0; row < sheet.RowCount; row++)
            bodyHeight += sheet.RowHeights.TryGetValue(row, out var height) ? height : sheet.DefaultRowHeight;
        var canvas = new Canvas
        {
            Width = bodyWidth,
            Height = bodyHeight,
            ClipToBounds = true,
            IsHitTestVisible = false
        };
        Grid.SetRow(canvas, 1);
        Grid.SetColumn(canvas, 1);
        Grid.SetRowSpan(canvas, sheet.RowCount);
        Grid.SetColumnSpan(canvas, sheet.ColumnCount);
        foreach (var image in sheet.Images)
        {
            var control = CreateImage(image);
            if (control is null) continue;
            Canvas.SetLeft(control, image.Left);
            Canvas.SetTop(control, image.Top);
            canvas.Children.Add(control);
        }
        grid.Children.Add(canvas);
    }

    private Image? CreateImage(XlsxImageModel image)
    {
        if (_workbook is null) return null;
        try
        {
            using var archive = _workbook.SourceBytes is { } sourceBytes
                ? new ZipArchive(new MemoryStream(sourceBytes, writable: false), ZipArchiveMode.Read)
                : new ZipArchive(new FileStream(_workbook.SourcePath ?? throw new InvalidOperationException("The workbook source is unavailable."), FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete), ZipArchiveMode.Read);
            var entry = archive.GetEntry(image.PackagePath);
            if (entry is null) return null;
            using var compressed = entry.Open();
            using var stream = new MemoryStream((int)Math.Min(entry.Length, int.MaxValue));
            compressed.CopyTo(stream);
            stream.Position = 0;
            var bitmap = Bitmap.DecodeToWidth(stream, Math.Clamp((int)Math.Ceiling(image.Width), 1, 1280), BitmapInterpolationMode.HighQuality);
            _ownedImages.Add(bitmap);
            return new Image { Source = bitmap, Width = image.Width, Height = image.Height, Stretch = Stretch.Uniform };
        }
        catch (Exception) { return null; }
    }

    private static string ColumnName(int zeroBased)
    {
        var result = string.Empty;
        for (var value = zeroBased + 1; value > 0; value = (value - 1) / 26) result = (char)('A' + (value - 1) % 26) + result;
        return result;
    }

    private void OnPointerWheelChanged(object? sender, PointerWheelEventArgs e)
    {
        if (!e.KeyModifiers.HasFlag(KeyModifiers.Control)) return;
        // Preserve normal scrolling; workbook zoom belongs to the host so it can be
        // shared with DOCX/PPTX instead of keeping independent visual scale state.
        e.Handled = true;
    }

    private void ShowError(Exception error)
    {
        _workbook = null;
        ClearVisuals();
        _sheetHost.Child = new TextBlock { Text = error.Message, Foreground = Brush.Parse("#9F1239"), TextWrapping = TextWrapping.Wrap, Margin = new Thickness(16) };
    }

    private void ClearVisuals()
    {
        _sheetHost.Child = null;
        ReleaseImages();
        _activeSheetIndex = -1;
        SheetCount = 0;
    }

    private void ReleaseImages()
    {
        foreach (var image in _ownedImages) image.Dispose();
        _ownedImages.Clear();
    }

    private static IBrush? ToBrush(string? color) => color is not null && Color.TryParse(color.StartsWith('#') ? color : "#" + color, out var parsed) ? new SolidColorBrush(parsed) : null;

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Clear();
    }
}
