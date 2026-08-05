using System;
using System.Collections.Generic;
using Avalonia;
using global::Avalonia.Layout;
using global::Avalonia.Media;

namespace OfficeViewer.Avalonia.Docx;

internal sealed class DocxDocumentModel
{
    public double PageWidth { get; set; } = 794;
    public double PageHeight { get; set; } = 1123;
    public Thickness PageMargin { get; set; } = new(96);
    public List<DocxBlock> Blocks { get; } = [];
    public Dictionary<string, byte[]> Images { get; } = new(StringComparer.Ordinal);
    public DocxNumbering Numbering { get; set; } = new();
}

internal abstract class DocxBlock;

internal sealed class DocxParagraph : DocxBlock
{
    public DocxParagraphStyle Style { get; init; } = new();
    public List<DocxInline> Inlines { get; } = [];
    public string? NumberingId { get; init; }
    public int NumberingLevel { get; init; }
}

internal sealed class DocxTable : DocxBlock
{
    public List<List<DocxTableCell>> Rows { get; } = [];
}

internal sealed class DocxTableCell
{
    public List<DocxBlock> Blocks { get; } = [];
    public int ColumnSpan { get; init; } = 1;
    public string? Background { get; init; }
}

internal abstract class DocxInline;

internal sealed class DocxTextRun : DocxInline
{
    public required string Text { get; init; }
    public DocxRunStyle Style { get; init; } = new();
}

internal sealed class DocxBreak : DocxInline
{
    public bool IsPageBreak { get; init; }
}

internal sealed class DocxTab : DocxInline;

internal sealed class DocxPicture : DocxInline
{
    public required string RelationshipId { get; init; }
    public double Width { get; init; }
    public double Height { get; init; }
    public bool IsFloating { get; init; }
    public double HorizontalOffset { get; init; }
    public double VerticalOffset { get; init; }
    public string? HorizontalRelativeTo { get; init; }
    public string? VerticalRelativeTo { get; init; }
}

internal sealed class DocxRunStyle
{
    public string? FontFamily { get; init; }
    public double? FontSize { get; init; }
    public FontWeight? FontWeight { get; init; }
    public FontStyle? FontStyle { get; init; }
    public bool? Underline { get; init; }
    public bool? StrikeThrough { get; init; }
    public string? Foreground { get; init; }
    public string? Highlight { get; init; }
    public BaselineAlignment? BaselineAlignment { get; init; }
    public double? CharacterSpacing { get; init; }

    public static DocxRunStyle Merge(DocxRunStyle? inherited, DocxRunStyle? local) => new()
    {
        FontFamily = local?.FontFamily ?? inherited?.FontFamily,
        FontSize = local?.FontSize ?? inherited?.FontSize,
        FontWeight = local?.FontWeight ?? inherited?.FontWeight,
        FontStyle = local?.FontStyle ?? inherited?.FontStyle,
        Underline = local?.Underline ?? inherited?.Underline,
        StrikeThrough = local?.StrikeThrough ?? inherited?.StrikeThrough,
        Foreground = local?.Foreground ?? inherited?.Foreground,
        Highlight = local?.Highlight ?? inherited?.Highlight,
        BaselineAlignment = local?.BaselineAlignment ?? inherited?.BaselineAlignment,
        CharacterSpacing = local?.CharacterSpacing ?? inherited?.CharacterSpacing
    };
}

internal sealed class DocxParagraphStyle
{
    public Thickness? Margin { get; init; }
    public Thickness? Padding { get; init; }
    public TextAlignment? TextAlignment { get; init; }
    public double? LineHeight { get; init; }
    public double? FirstLineIndent { get; init; }
    public double? FirstTabStop { get; init; }
    public string? Background { get; init; }
    public DocxRunStyle RunStyle { get; init; } = new();

    public static DocxParagraphStyle Merge(DocxParagraphStyle? inherited, DocxParagraphStyle? local) => new()
    {
        Margin = local?.Margin ?? inherited?.Margin,
        Padding = local?.Padding ?? inherited?.Padding,
        TextAlignment = local?.TextAlignment ?? inherited?.TextAlignment,
        LineHeight = local?.LineHeight ?? inherited?.LineHeight,
        FirstLineIndent = local?.FirstLineIndent ?? inherited?.FirstLineIndent,
        FirstTabStop = local?.FirstTabStop ?? inherited?.FirstTabStop,
        Background = local?.Background ?? inherited?.Background,
        RunStyle = DocxRunStyle.Merge(inherited?.RunStyle, local?.RunStyle)
    };
}

internal sealed class DocxNumbering
{
    public Dictionary<string, Dictionary<int, DocxNumberingLevel>> Definitions { get; } = new(StringComparer.Ordinal);

    public sealed class DocxNumberingLevel
    {
        public int Start { get; init; } = 1;
        public string Format { get; init; } = "decimal";
        public string Text { get; init; } = "%1.";
        public Thickness? Indent { get; init; }
        public double? FirstLineIndent { get; init; }
        public string LabelAlignment { get; init; } = "left";
        public DocxRunStyle RunStyle { get; init; } = new();
    }
}
