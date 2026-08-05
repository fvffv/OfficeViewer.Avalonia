using System.Collections.Generic;
using Avalonia;
using global::Avalonia.Layout;
using global::Avalonia.Media;

namespace OfficeViewer.Avalonia.Xlsx;

internal sealed class XlsxWorkbookModel
{
    // Kept only while Document is non-null so sheet images can remain lazily decoded.
    public string? SourcePath { get; init; }
    public byte[]? SourceBytes { get; init; }
    public List<XlsxSheetModel> Sheets { get; } = [];
}

internal sealed class XlsxSheetModel
{
    public required string Name { get; init; }
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
    public double DefaultRowHeight { get; init; } = 24;
    public double DefaultColumnWidth { get; init; } = 80;
    public Dictionary<int, double> RowHeights { get; } = [];
    public Dictionary<int, double> ColumnWidths { get; } = [];
    public Dictionary<(int Row, int Column), XlsxCellModel> Cells { get; } = [];
    public List<XlsxImageModel> Images { get; } = [];
}

internal sealed class XlsxImageModel
{
    public required string PackagePath { get; init; }
    public double Left { get; init; }
    public double Top { get; init; }
    public double Width { get; init; }
    public double Height { get; init; }
}

internal sealed class XlsxCellModel
{
    public required int Row { get; init; }
    public required int Column { get; init; }
    public string Text { get; init; } = string.Empty;
    public List<XlsxTextRun> Runs { get; } = [];
    public XlsxCellStyle Style { get; init; } = XlsxCellStyle.Default;
    public int RowSpan { get; set; } = 1;
    public int ColumnSpan { get; set; } = 1;
    public bool IsMergedChild { get; set; }
}

internal sealed class XlsxTextRun
{
    public required string Text { get; init; }
    public string? Foreground { get; init; }
    public double? FontSize { get; init; }
    public string? FontFamily { get; init; }
    public FontWeight? FontWeight { get; init; }
    public FontStyle? FontStyle { get; init; }
    public bool Underline { get; init; }
}

internal sealed class XlsxCellStyle
{
    public static XlsxCellStyle Default { get; } = new();
    public string? Background { get; init; }
    public string? Foreground { get; init; }
    public string? FontFamily { get; init; }
    public double? FontSize { get; init; }
    public FontWeight? FontWeight { get; init; }
    public FontStyle? FontStyle { get; init; }
    public bool Underline { get; init; }
    public bool Wrap { get; init; }
    public TextAlignment HorizontalAlignment { get; init; } = TextAlignment.Left;
    public VerticalAlignment VerticalAlignment { get; init; } = VerticalAlignment.Center;
    public Thickness BorderThickness { get; init; }
    public string? BorderColor { get; init; }
    public string? NumberFormat { get; init; }
}
