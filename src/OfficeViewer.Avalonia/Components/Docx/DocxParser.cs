using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Avalonia;
using global::Avalonia.Layout;
using global::Avalonia.Media;

namespace OfficeViewer.Avalonia.Docx;

/// <summary>
/// Reads the same WordprocessingML primitives consumed by vue-office's docx-preview renderer,
/// but produces platform-neutral models for Avalonia controls rather than HTML/CSS nodes.
/// </summary>
internal sealed class DocxParser
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";

    private readonly Dictionary<string, StyleDefinition> _styles = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _relationships = new(StringComparer.Ordinal);
    private readonly DocxNumbering _numbering = new();
    private string? _defaultParagraphStyleId;

    public DocxDocumentModel Parse(string path)
    {
        using var stream = File.OpenRead(path);
        return Parse(stream);
    }

    public DocxDocumentModel Parse(Stream source)
    {
        using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: false);
        var document = new DocxDocumentModel();
        ReadRelationships(archive);
        ReadStyles(archive);
        ReadNumbering(archive);
        document.Numbering = _numbering;
        ReadImages(archive, document);

        var documentXml = ReadXml(archive, "word/document.xml");
        var body = documentXml.Root?.Element(W + "body")
            ?? throw new InvalidDataException("The DOCX package has no word/document.xml body.");

        var page = body.Elements(W + "sectPr").LastOrDefault()
            ?? body.Descendants(W + "sectPr").LastOrDefault();
        if (page is not null)
        {
            document.PageWidth = TwipsToPixels(IntAttribute(page.Element(W + "pgSz"), "w", 11910));
            document.PageHeight = TwipsToPixels(IntAttribute(page.Element(W + "pgSz"), "h", 16840));
            document.PageMargin = ReadPageMargin(page.Element(W + "pgMar"));
        }

        foreach (var element in body.Elements())
        {
            if (element.Name == W + "p")
                document.Blocks.Add(ParseParagraph(element));
            else if (element.Name == W + "tbl")
                document.Blocks.Add(ParseTable(element));
        }

        return document;
    }

    private static XDocument ReadXml(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name) ?? throw new InvalidDataException($"The DOCX package is missing {name}.");
        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private void ReadRelationships(ZipArchive archive)
    {
        var entry = archive.GetEntry("word/_rels/document.xml.rels");
        if (entry is null) return;
        using var stream = entry.Open();
        var xml = XDocument.Load(stream);
        foreach (var relationship in xml.Root?.Elements() ?? [])
        {
            var id = (string?)relationship.Attribute("Id");
            var target = (string?)relationship.Attribute("Target");
            if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(target))
                _relationships[id] = target;
        }
    }

    private void ReadImages(ZipArchive archive, DocxDocumentModel document)
    {
        foreach (var (relationshipId, target) in _relationships)
        {
            var packagePath = ResolveWordTarget(target);
            if (!packagePath.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase)) continue;
            var entry = archive.GetEntry(packagePath);
            if (entry is null) continue;
            using var input = entry.Open();
            using var output = new MemoryStream();
            input.CopyTo(output);
            document.Images[relationshipId] = output.ToArray();
        }
    }

    private void ReadStyles(ZipArchive archive)
    {
        var entry = archive.GetEntry("word/styles.xml");
        if (entry is null) return;
        using var stream = entry.Open();
        var xml = XDocument.Load(stream);
        foreach (var style in xml.Descendants(W + "style"))
        {
            var id = Value(style, "styleId");
            if (string.IsNullOrWhiteSpace(id)) continue;
            if (Value(style, "type") == "paragraph" && Value(style, "default") is "1" or "true")
                _defaultParagraphStyleId = id;
            _styles[id] = new StyleDefinition
            {
                BasedOn = Value(style.Element(W + "basedOn"), "val"),
                Paragraph = ReadParagraphStyle(style.Element(W + "pPr")),
                Run = ReadRunStyle(style.Element(W + "rPr"))
            };
        }
    }

    private void ReadNumbering(ZipArchive archive)
    {
        var entry = archive.GetEntry("word/numbering.xml");
        if (entry is null) return;
        using var stream = entry.Open();
        var xml = XDocument.Load(stream);
        var abstractNumbers = xml.Descendants(W + "abstractNum")
            .ToDictionary(x => Value(x, "abstractNumId") ?? string.Empty, StringComparer.Ordinal);

        foreach (var number in xml.Descendants(W + "num"))
        {
            var numberId = Value(number, "numId");
            var abstractId = Value(number.Element(W + "abstractNumId"), "val");
            if (string.IsNullOrWhiteSpace(numberId) || string.IsNullOrWhiteSpace(abstractId) ||
                !abstractNumbers.TryGetValue(abstractId, out var abstractNumber)) continue;

            var levels = new Dictionary<int, DocxNumbering.DocxNumberingLevel>();
            foreach (var level in abstractNumber.Elements(W + "lvl"))
            {
                var levelId = IntAttribute(level, "ilvl", 0);
                var paragraphProperties = level.Element(W + "pPr");
                levels[levelId] = new DocxNumbering.DocxNumberingLevel
                {
                    Start = IntAttribute(level.Element(W + "start"), "val", 1),
                    Format = Value(level.Element(W + "numFmt"), "val") ?? "decimal",
                    Text = Value(level.Element(W + "lvlText"), "val") ?? "%1.",
                    Indent = ReadIndent(paragraphProperties?.Element(W + "ind")),
                    FirstLineIndent = ReadFirstLineIndent(paragraphProperties?.Element(W + "ind")),
                    LabelAlignment = Value(level.Element(W + "lvlJc"), "val") ?? "left",
                    RunStyle = ReadRunStyle(level.Element(W + "rPr"))
                };
            }
            _numbering.Definitions[numberId] = levels;
        }
    }

    private DocxParagraph ParseParagraph(XElement paragraph)
    {
        var pPr = paragraph.Element(W + "pPr");
        var styleId = Value(pPr?.Element(W + "pStyle"), "val");
        // Word applies the document's default paragraph style when w:pStyle is omitted.
        // Without this fallback, runs in otherwise identical paragraphs inherit Avalonia's
        // system font rather than the DOCX's declared default font.
        var style = ResolveParagraphStyle(styleId ?? _defaultParagraphStyleId);
        style = DocxParagraphStyle.Merge(style, ReadParagraphStyle(pPr));
        var numberProperties = pPr?.Element(W + "numPr");
        var result = new DocxParagraph
        {
            Style = style,
            NumberingId = Value(numberProperties?.Element(W + "numId"), "val"),
            NumberingLevel = IntAttribute(numberProperties?.Element(W + "ilvl"), "val", 0)
        };

        foreach (var child in paragraph.Elements().Where(x => x.Name != W + "pPr"))
            ParseInlineContainer(child, style.RunStyle, result.Inlines);
        return result;
    }

    private DocxTable ParseTable(XElement table)
    {
        var result = new DocxTable();
        foreach (var row in table.Elements(W + "tr"))
        {
            var parsedRow = new List<DocxTableCell>();
            foreach (var cell in row.Elements(W + "tc"))
            {
                var cellProperties = cell.Element(W + "tcPr");
                var parsedCell = new DocxTableCell
                {
                    ColumnSpan = Math.Max(1, IntAttribute(cellProperties?.Element(W + "gridSpan"), "val", 1)),
                    Background = Value(cellProperties?.Element(W + "shd"), "fill")
                };
                foreach (var block in cell.Elements().Where(x => x.Name == W + "p" || x.Name == W + "tbl"))
                    parsedCell.Blocks.Add(block.Name == W + "p" ? ParseParagraph(block) : ParseTable(block));
                parsedRow.Add(parsedCell);
            }
            result.Rows.Add(parsedRow);
        }
        return result;
    }

    private void ParseInlineContainer(XElement element, DocxRunStyle inherited, List<DocxInline> output)
    {
        if (element.Name == W + "r")
        {
            var style = DocxRunStyle.Merge(inherited, ReadRunStyle(element.Element(W + "rPr")));
            foreach (var child in element.Elements().Where(x => x.Name != W + "rPr"))
            {
                if (child.Name == W + "t" || child.Name == W + "delText")
                    output.Add(new DocxTextRun { Text = child.Value, Style = style });
                else if (child.Name == W + "tab")
                    output.Add(new DocxTab());
                else if (child.Name == W + "br" || child.Name == W + "cr" || child.Name == W + "lastRenderedPageBreak")
                    output.Add(new DocxBreak { IsPageBreak = Value(child, "type") == "page" || child.Name == W + "lastRenderedPageBreak" });
                else if (child.Name == W + "drawing" || child.Name == W + "pict")
                    ParsePictures(child, output);
            }
            return;
        }

        if (element.Name == W + "hyperlink" || element.Name == W + "smartTag" || element.Name == W + "sdt" || element.Name == W + "sdtContent" || element.Name == W + "ins")
        {
            foreach (var child in element.Elements().Where(x => x.Name != W + "sdtPr"))
                ParseInlineContainer(child, inherited, output);
        }
    }

    private void ParsePictures(XElement drawing, List<DocxInline> output)
    {
        var floating = drawing.Descendants(Wp + "anchor").FirstOrDefault();
        var inline = drawing.Descendants(Wp + "inline").FirstOrDefault();
        var shape = floating ?? inline;
        var extent = shape?.Element(Wp + "extent") ?? drawing.Descendants(A + "ext").FirstOrDefault();
        var width = EmuToPixels(LongAttribute(extent, "cx", 1_905_000));
        var height = EmuToPixels(LongAttribute(extent, "cy", 1_428_750));
        var positionH = floating?.Element(Wp + "positionH");
        var positionV = floating?.Element(Wp + "positionV");
        var horizontalOffset = EmuToPixels(LongValue(positionH?.Element(Wp + "posOffset")));
        var verticalOffset = EmuToPixels(LongValue(positionV?.Element(Wp + "posOffset")));

        foreach (var blip in drawing.Descendants(A + "blip"))
        {
            var relationshipId = (string?)blip.Attribute(R + "embed");
            if (!string.IsNullOrWhiteSpace(relationshipId))
                output.Add(new DocxPicture
                {
                    RelationshipId = relationshipId,
                    Width = width,
                    Height = height,
                    IsFloating = floating is not null,
                    HorizontalOffset = horizontalOffset,
                    VerticalOffset = verticalOffset,
                    HorizontalRelativeTo = (string?)positionH?.Attribute("relativeFrom"),
                    VerticalRelativeTo = (string?)positionV?.Attribute("relativeFrom")
                });
        }

        foreach (var imageData in drawing.Descendants(V + "imagedata"))
        {
            var relationshipId = (string?)imageData.Attribute(R + "id");
            if (!string.IsNullOrWhiteSpace(relationshipId))
                output.Add(new DocxPicture { RelationshipId = relationshipId, Width = width, Height = height, IsFloating = floating is not null });
        }
    }

    private DocxParagraphStyle ResolveParagraphStyle(string? styleId)
    {
        var chain = ResolveStyleChain(styleId);
        var result = new DocxParagraphStyle();
        foreach (var definition in chain)
            result = DocxParagraphStyle.Merge(result, DocxParagraphStyle.Merge(definition.Paragraph, new DocxParagraphStyle { RunStyle = definition.Run }));
        return result;
    }

    private List<StyleDefinition> ResolveStyleChain(string? styleId)
    {
        var result = new List<StyleDefinition>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        while (!string.IsNullOrWhiteSpace(styleId) && seen.Add(styleId) && _styles.TryGetValue(styleId, out var definition))
        {
            result.Insert(0, definition);
            styleId = definition.BasedOn;
        }
        return result;
    }

    private static DocxRunStyle ReadRunStyle(XElement? properties) => new()
    {
        FontFamily = Value(properties?.Element(W + "rFonts"), "eastAsia") ?? Value(properties?.Element(W + "rFonts"), "ascii") ?? Value(properties?.Element(W + "rFonts"), "hAnsi"),
        FontSize = IntAttribute(properties?.Element(W + "sz"), "val", 0) is var halfPoints && halfPoints > 0 ? HalfPointsToPixels(halfPoints) : null,
        FontWeight = properties?.Element(W + "b") is not null ? FontWeight.Bold : null,
        FontStyle = properties?.Element(W + "i") is not null ? FontStyle.Italic : null,
        Underline = properties?.Element(W + "u") is { } underline ? Value(underline, "val") != "none" : null,
        StrikeThrough = properties?.Element(W + "strike") is not null || properties?.Element(W + "dstrike") is not null ? true : null,
        Foreground = NormalizeColor(Value(properties?.Element(W + "color"), "val")),
        Highlight = NormalizeHighlight(Value(properties?.Element(W + "highlight"), "val")),
        BaselineAlignment = ParseBaseline(Value(properties?.Element(W + "vertAlign"), "val")),
        CharacterSpacing = TwipsToPixels(IntAttribute(properties?.Element(W + "spacing"), "val", 0))
    };

    private static DocxParagraphStyle ReadParagraphStyle(XElement? properties)
    {
        var spacing = properties?.Element(W + "spacing");
        var indent = ReadIndent(properties?.Element(W + "ind"));
        var before = TwipsToPixels(IntAttribute(spacing, "before", 0));
        var after = TwipsToPixels(IntAttribute(spacing, "after", 0));
        var line = IntAttribute(spacing, "line", 0);
        var lineRule = Value(spacing, "lineRule");
        // Word's automatic line spacing is a multiplier, not a fixed 16px-based height.
        // Let Avalonia measure it from the largest run, otherwise large text is clipped and
        // the following paragraph can overlap it.
        double? lineHeight = line == 0 || lineRule is null or "auto" ? null : TwipsToPixels(line);
        return new DocxParagraphStyle
        {
            Margin = new Thickness(indent?.Left ?? 0, before, indent?.Right ?? 0, after),
            Padding = null,
            TextAlignment = ParseAlignment(Value(properties?.Element(W + "jc"), "val")),
            LineHeight = lineHeight,
            FirstLineIndent = ReadFirstLineIndent(properties?.Element(W + "ind")),
            FirstTabStop = ReadFirstTabStop(properties?.Element(W + "tabs")),
            Background = NormalizeColor(Value(properties?.Element(W + "shd"), "fill")),
            RunStyle = ReadRunStyle(properties?.Element(W + "rPr"))
        };
    }

    private static Thickness? ReadIndent(XElement? indent)
    {
        if (indent is null) return null;
        return new Thickness(
            TwipsToPixels(IntAttribute(indent, "left", IntAttribute(indent, "start", 0))),
            0,
            TwipsToPixels(IntAttribute(indent, "right", IntAttribute(indent, "end", 0))),
            0);
    }

    private static double? ReadFirstLineIndent(XElement? indent)
    {
        if (indent is null) return null;
        var firstLine = IntAttribute(indent, "firstLine", 0);
        if (firstLine != 0) return TwipsToPixels(firstLine);
        var hanging = IntAttribute(indent, "hanging", 0);
        return hanging == 0 ? null : -TwipsToPixels(hanging);
    }

    private static double? ReadFirstTabStop(XElement? tabs)
    {
        var tab = tabs?.Elements(W + "tab")
            .FirstOrDefault(x => Value(x, "val") is "left" or "start");
        return tab is null ? null : TwipsToPixels(IntAttribute(tab, "pos", 0));
    }

    private static Thickness ReadPageMargin(XElement? margin) => new(
        TwipsToPixels(IntAttribute(margin, "left", 1440)),
        TwipsToPixels(IntAttribute(margin, "top", 1440)),
        TwipsToPixels(IntAttribute(margin, "right", 1440)),
        TwipsToPixels(IntAttribute(margin, "bottom", 1440)));

    private static TextAlignment? ParseAlignment(string? value) => value switch
    {
        "center" => TextAlignment.Center,
        "right" or "end" => TextAlignment.Right,
        "both" or "distribute" => TextAlignment.Justify,
        "left" or "start" => TextAlignment.Left,
        _ => null
    };

    private static BaselineAlignment? ParseBaseline(string? value) => value switch
    {
        "superscript" => BaselineAlignment.TextTop,
        "subscript" => BaselineAlignment.Subscript,
        _ => null
    };

    private static string? NormalizeColor(string? value) => string.IsNullOrWhiteSpace(value) || value is "auto" or "none" ? null : value.TrimStart('#');
    private static string? NormalizeHighlight(string? value) => value switch
    {
        null or "none" => null,
        "yellow" => "FFFF00",
        "green" => "00FF00",
        "cyan" => "00FFFF",
        "magenta" => "FF00FF",
        "blue" => "0000FF",
        "red" => "FF0000",
        "darkBlue" => "000080",
        "darkRed" => "800000",
        "darkGreen" => "008000",
        "darkYellow" => "808000",
        "darkCyan" => "008080",
        "darkMagenta" => "800080",
        "gray" => "808080",
        "lightGray" => "C0C0C0",
        "black" => "000000",
        _ => value
    };

    private static string ResolveWordTarget(string target) => "word/" + target.Replace('\\', '/').TrimStart('/');
    private static string? Value(XElement? element, string name) => (string?)element?.Attribute(W + name);
    private static int IntAttribute(XElement? element, string name, int fallback) => int.TryParse(Value(element, name), out var value) ? value : fallback;
    private static long LongAttribute(XElement? element, string name, long fallback) => long.TryParse(Value(element, name), out var value) ? value : fallback;
    private static long LongValue(XElement? element) => long.TryParse(element?.Value, out var value) ? value : 0;
    private static double TwipsToPixels(int twips) => twips * 96d / 1440d;
    private static double HalfPointsToPixels(int halfPoints) => halfPoints <= 0 ? 0 : halfPoints * 96d / 144d;
    private static double EmuToPixels(long emu) => emu * 96d / 914400d;

    private sealed class StyleDefinition
    {
        public string? BasedOn { get; init; }
        public DocxParagraphStyle Paragraph { get; init; } = new();
        public DocxRunStyle Run { get; init; } = new();
    }
}
