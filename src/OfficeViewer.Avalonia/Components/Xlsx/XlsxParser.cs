using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Avalonia;
using global::Avalonia.Layout;
using global::Avalonia.Media;

namespace OfficeViewer.Avalonia.Xlsx;

/// <summary>
/// OpenXML workbook reader modelled on vue-office's ExcelJS transfer: workbook
/// relationship order, shared/rich strings, row/column dimensions, styles,
/// merges, cached formula values and hyperlinks are transferred into a compact
/// native model without ExcelJS or a JavaScript runtime.
/// </summary>
internal sealed class XlsxParser
{
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private readonly List<XlsxCellStyle> _styles = [];
    private readonly List<List<XlsxTextRun>> _sharedStrings = [];
    private readonly Dictionary<int, string> _numberFormats = new();
    private readonly Dictionary<int, string> _themeColors = new();

    public XlsxWorkbookModel Parse(string path)
    {
        // Excel/WPS often holds an open workbook with write sharing enabled.
        // The viewer only reads the package and must not reject that normal case.
        using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return Parse(source, path, null);
    }

    public XlsxWorkbookModel Parse(byte[] document)
    {
        ArgumentNullException.ThrowIfNull(document);
        using var source = new MemoryStream(document, writable: false);
        return Parse(source, null, document);
    }

    private XlsxWorkbookModel Parse(Stream source, string? sourcePath, byte[]? sourceBytes)
    {
        using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: false);
        ReadTheme(archive);
        ReadStyles(archive);
        ReadSharedStrings(archive);
        var workbook = ReadXml(archive, "xl/workbook.xml");
        var relations = ReadRelationships(archive, "xl/workbook.xml");
        var result = new XlsxWorkbookModel { SourcePath = sourcePath, SourceBytes = sourceBytes };
        foreach (var item in workbook.Descendants(S + "sheet"))
        {
            var relationshipId = (string?)item.Attribute(R + "id");
            if (relationshipId is null || !relations.TryGetValue(relationshipId, out var sheetPath)) continue;
            result.Sheets.Add(ParseSheet(archive, sheetPath, (string?)item.Attribute("name") ?? "Sheet"));
        }
        return result;
    }

    private XlsxSheetModel ParseSheet(ZipArchive archive, string sheetPath, string name)
    {
        var xml = ReadXml(archive, sheetPath);
        var format = xml.Root?.Element(S + "sheetFormatPr");
        var sheet = new XlsxSheetModel
        {
            Name = name,
            // SpreadsheetML row heights use typographic points. Avalonia layout uses
            // 96-DPI device-independent pixels, so retaining the raw number makes
            // DrawingML pictures (which are already converted from EMU to pixels)
            // overflow their anchored cells.
            DefaultRowHeight = Math.Max(1, PointsToDips(DoubleAttribute(format, "defaultRowHeight", 18))),
            DefaultColumnWidth = Math.Max(40, DoubleAttribute(format, "defaultColWidth", 8.43) * 6)
        };
        foreach (var column in xml.Root?.Element(S + "cols")?.Elements(S + "col") ?? [])
        {
            var min = IntAttribute(column, "min", 1) - 1;
            var max = IntAttribute(column, "max", min + 1) - 1;
            var width = AttributeTrue(column, "hidden") ? 0 : Math.Max(1, DoubleAttribute(column, "width", sheet.DefaultColumnWidth / 6) * 6);
            for (var index = Math.Max(0, min); index <= max; index++) sheet.ColumnWidths[index] = width;
            sheet.ColumnCount = Math.Max(sheet.ColumnCount, max + 1);
        }

        foreach (var row in xml.Descendants(S + "row"))
        {
            var rowIndex = Math.Max(0, IntAttribute(row, "r", 1) - 1);
            if (AttributeTrue(row, "hidden")) sheet.RowHeights[rowIndex] = 0;
            else if (row.Attribute("ht") is not null) sheet.RowHeights[rowIndex] = Math.Max(1, PointsToDips(DoubleAttribute(row, "ht", 18)));
            sheet.RowCount = Math.Max(sheet.RowCount, rowIndex + 1);
            foreach (var cell in row.Elements(S + "c"))
            {
                if (!TryParseAddress((string?)cell.Attribute("r"), out var column, out var parsedRow)) continue;
                rowIndex = parsedRow;
                var styleIndex = IntAttribute(cell, "s", 0);
                var style = styleIndex >= 0 && styleIndex < _styles.Count ? _styles[styleIndex] : XlsxCellStyle.Default;
                var (text, runs) = ReadCellValue(cell, style);
                var model = new XlsxCellModel { Row = rowIndex, Column = column, Text = text, Style = style };
                model.Runs.AddRange(runs);
                sheet.Cells[(rowIndex, column)] = model;
                sheet.RowCount = Math.Max(sheet.RowCount, rowIndex + 1);
                sheet.ColumnCount = Math.Max(sheet.ColumnCount, column + 1);
            }
        }

        foreach (var merge in xml.Descendants(S + "mergeCell")) ApplyMerge(sheet, (string?)merge.Attribute("ref"));
        var sheetRelationships = ReadRelationships(archive, sheetPath);
        var drawingId = (string?)xml.Root?.Element(S + "drawing")?.Attribute(R + "id");
        if (drawingId is not null && sheetRelationships.TryGetValue(drawingId, out var drawingPath))
            ReadDrawingImages(archive, drawingPath, sheet);
        // Vue-office always keeps a usable grid even for sparse worksheets.
        sheet.RowCount = Math.Max(sheet.RowCount, 1);
        sheet.ColumnCount = Math.Max(sheet.ColumnCount, 1);
        return sheet;
    }

    private static void ReadDrawingImages(ZipArchive archive, string drawingPath, XlsxSheetModel sheet)
    {
        var drawing = ReadXml(archive, drawingPath);
        var relationships = ReadRelationships(archive, drawingPath);
        foreach (var anchor in drawing.Root?.Elements().Where(x => x.Name == Xdr + "twoCellAnchor" || x.Name == Xdr + "oneCellAnchor") ?? [])
        {
            var picture = anchor.Element(Xdr + "pic");
            var relationshipId = (string?)picture?.Descendants(A + "blip").FirstOrDefault()?.Attribute(R + "embed");
            if (relationshipId is null || !relationships.TryGetValue(relationshipId, out var packagePath)) continue;
            var from = anchor.Element(Xdr + "from");
            var to = anchor.Element(Xdr + "to");
            var fromColumn = IntElement(from, Xdr + "col", 0); var fromRow = IntElement(from, Xdr + "row", 0);
            var toColumn = IntElement(to, Xdr + "col", fromColumn); var toRow = IntElement(to, Xdr + "row", fromRow);
            var left = OffsetToPixels(sheet, fromColumn, fromRow, LongElement(from, Xdr + "colOff", 0), LongElement(from, Xdr + "rowOff", 0));
            var right = OffsetToPixels(sheet, toColumn, toRow, LongElement(to, Xdr + "colOff", 0), LongElement(to, Xdr + "rowOff", 0));
            if (to is null)
            {
                var extent = picture?.Element(Xdr + "spPr")?.Element(A + "xfrm")?.Element(A + "ext");
                right = new Point(left.X + LongAttribute(extent, "cx", 0) * 96d / 914400d, left.Y + LongAttribute(extent, "cy", 0) * 96d / 914400d);
            }
            sheet.Images.Add(new XlsxImageModel { PackagePath = packagePath, Left = left.X, Top = left.Y, Width = Math.Max(1, right.X - left.X), Height = Math.Max(1, right.Y - left.Y) });
        }
    }

    private static Point OffsetToPixels(XlsxSheetModel sheet, int column, int row, long columnOffset, long rowOffset)
    {
        var x = Enumerable.Range(0, Math.Max(0, column)).Sum(index => sheet.ColumnWidths.TryGetValue(index, out var width) ? width : sheet.DefaultColumnWidth);
        var y = Enumerable.Range(0, Math.Max(0, row)).Sum(index => sheet.RowHeights.TryGetValue(index, out var height) ? height : sheet.DefaultRowHeight);
        return new Point(x + columnOffset * 96d / 914400d, y + rowOffset * 96d / 914400d);
    }

    private static double PointsToDips(double points) => points * 96d / 72d;

    private (string Text, List<XlsxTextRun> Runs) ReadCellValue(XElement cell, XlsxCellStyle style)
    {
        var type = (string?)cell.Attribute("t");
        var raw = cell.Element(S + "v")?.Value ?? string.Empty;
        if (type == "s" && int.TryParse(raw, out var sharedIndex) && sharedIndex >= 0 && sharedIndex < _sharedStrings.Count)
        {
            var runs = _sharedStrings[sharedIndex].Select(CloneRun).ToList();
            return (string.Concat(runs.Select(x => x.Text)), runs);
        }
        if (type == "inlineStr")
        {
            var runs = ReadRichText(cell.Element(S + "is"));
            return (string.Concat(runs.Select(x => x.Text)), runs);
        }
        if (type == "b") return (raw == "1" ? "TRUE" : "FALSE", []);
        if (type == "e") return (raw, []);
        if (string.IsNullOrEmpty(raw)) return (string.Empty, []);
        return (FormatValue(raw, style.NumberFormat), []);
    }

    private void ReadSharedStrings(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry is null) return;
        using var stream = entry.Open();
        var xml = XDocument.Load(stream);
        foreach (var item in xml.Descendants(S + "si")) _sharedStrings.Add(ReadRichText(item));
    }

    private List<XlsxTextRun> ReadRichText(XElement? source)
    {
        var result = new List<XlsxTextRun>();
        if (source is null) return result;
        var runs = source.Elements(S + "r").ToList();
        if (runs.Count == 0)
        {
            var text = string.Concat(source.Descendants(S + "t").Select(x => x.Value));
            if (!string.IsNullOrEmpty(text)) result.Add(new XlsxTextRun { Text = text });
            return result;
        }
        foreach (var run in runs)
        {
            var properties = run.Element(S + "rPr");
            var fontColor = ResolveColor(properties?.Element(S + "color"));
            result.Add(new XlsxTextRun
            {
                Text = string.Concat(run.Descendants(S + "t").Select(x => x.Value)),
                Foreground = fontColor,
                FontSize = DoubleAttribute(properties?.Element(S + "sz"), "val", 0) is var size && size > 0 ? size * 96d / 72d : null,
                FontFamily = (string?)properties?.Element(S + "rFont")?.Attribute("val"),
                FontWeight = properties?.Element(S + "b") is not null ? FontWeight.Bold : null,
                FontStyle = properties?.Element(S + "i") is not null ? FontStyle.Italic : null,
                Underline = properties?.Element(S + "u") is not null
            });
        }
        return result;
    }

    private void ReadStyles(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/styles.xml");
        if (entry is null) { _styles.Add(XlsxCellStyle.Default); return; }
        using var stream = entry.Open();
        var xml = XDocument.Load(stream);
        foreach (var format in xml.Descendants(S + "numFmt"))
            _numberFormats[IntAttribute(format, "numFmtId", 0)] = (string?)format.Attribute("formatCode") ?? string.Empty;
        var fonts = xml.Root?.Element(S + "fonts")?.Elements(S + "font").Select(ReadFont).ToList() ?? [];
        var fills = xml.Root?.Element(S + "fills")?.Elements(S + "fill").Select(ReadFill).ToList() ?? [];
        var borders = xml.Root?.Element(S + "borders")?.Elements(S + "border").Select(ReadBorder).ToList() ?? [];
        foreach (var xf in xml.Root?.Element(S + "cellXfs")?.Elements(S + "xf") ?? [])
        {
            var font = AtFont(fonts, IntAttribute(xf, "fontId", 0));
            var fill = AtString(fills, IntAttribute(xf, "fillId", 0));
            var border = AtBorder(borders, IntAttribute(xf, "borderId", 0));
            var alignment = xf.Element(S + "alignment");
            var numFmtId = IntAttribute(xf, "numFmtId", 0);
            _styles.Add(new XlsxCellStyle
            {
                Background = fill,
                Foreground = font.Color,
                FontFamily = font.Family,
                FontSize = font.Size,
                FontWeight = font.Bold ? FontWeight.Bold : null,
                FontStyle = font.Italic ? FontStyle.Italic : null,
                Underline = font.Underline,
                Wrap = AttributeTrue(alignment, "wrapText"),
                HorizontalAlignment = ParseHorizontal((string?)alignment?.Attribute("horizontal")),
                VerticalAlignment = ParseVertical((string?)alignment?.Attribute("vertical")),
                BorderThickness = border.Thickness,
                BorderColor = border.Color,
                NumberFormat = _numberFormats.TryGetValue(numFmtId, out var custom) ? custom : BuiltinNumberFormat(numFmtId)
            });
        }
        if (_styles.Count == 0) _styles.Add(XlsxCellStyle.Default);
    }

    private void ReadTheme(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/theme/theme1.xml");
        if (entry is null) return;
        using var stream = entry.Open();
        var xml = XDocument.Load(stream);
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var scheme = xml.Descendants(a + "clrScheme").FirstOrDefault();
        if (scheme is null) return;
        var indexes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["lt1"] = 0, ["dk1"] = 1, ["lt2"] = 2, ["dk2"] = 3,
            ["accent1"] = 4, ["accent2"] = 5, ["accent3"] = 6, ["accent4"] = 7,
            ["accent5"] = 8, ["accent6"] = 9, ["hlink"] = 10, ["folHlink"] = 11
        };
        foreach (var item in scheme.Elements())
        {
            if (!indexes.TryGetValue(item.Name.LocalName, out var index)) continue;
            var value = item.Descendants(a + "srgbClr").Attributes("val").FirstOrDefault()?.Value
                ?? item.Descendants(a + "sysClr").Attributes("lastClr").FirstOrDefault()?.Value;
            if (value is not null) _themeColors[index] = value;
        }
    }

    private FontData ReadFont(XElement font) => new(
        (string?)font.Element(S + "name")?.Attribute("val"),
        DoubleAttribute(font.Element(S + "sz"), "val", 11) * 96d / 72d,
        ResolveColor(font.Element(S + "color")), font.Element(S + "b") is not null,
        font.Element(S + "i") is not null, font.Element(S + "u") is not null);

    private string? ReadFill(XElement fill)
    {
        var pattern = fill.Element(S + "patternFill");
        return (string?)pattern?.Attribute("patternType") == "solid" ? ResolveColor(pattern?.Element(S + "fgColor")) : null;
    }

    private BorderData ReadBorder(XElement border)
    {
        var left = border.Element(S + "left"); var right = border.Element(S + "right");
        var top = border.Element(S + "top"); var bottom = border.Element(S + "bottom");
        var color = ResolveColor(left?.Element(S + "color")) ?? ResolveColor(right?.Element(S + "color")) ?? ResolveColor(top?.Element(S + "color")) ?? ResolveColor(bottom?.Element(S + "color"));
        return new BorderData(new Thickness(BorderWidth(left), BorderWidth(top), BorderWidth(right), BorderWidth(bottom)), color);
    }

    private string? ResolveColor(XElement? color)
    {
        if (color is null) return null;
        var rgb = (string?)color.Attribute("rgb");
        if (!string.IsNullOrWhiteSpace(rgb)) return rgb;
        if (int.TryParse((string?)color.Attribute("theme"), out var theme) && _themeColors.TryGetValue(theme, out var themeColor))
            return ApplyTint(themeColor, DoubleAttribute(color, "tint", 0));
        if (int.TryParse((string?)color.Attribute("indexed"), out var indexed)) return IndexedColor(indexed);
        return AttributeTrue(color, "auto") ? "000000" : null;
    }

    private static void ApplyMerge(XlsxSheetModel sheet, string? range)
    {
        if (string.IsNullOrWhiteSpace(range)) return;
        var parts = range.Split(':');
        if (!TryParseAddress(parts[0], out var startColumn, out var startRow) || !TryParseAddress(parts[^1], out var endColumn, out var endRow)) return;
        var master = GetOrCreateCell(sheet, startRow, startColumn);
        master.RowSpan = Math.Max(1, endRow - startRow + 1);
        master.ColumnSpan = Math.Max(1, endColumn - startColumn + 1);
        for (var row = startRow; row <= endRow; row++)
            for (var column = startColumn; column <= endColumn; column++)
                if (row != startRow || column != startColumn) GetOrCreateCell(sheet, row, column).IsMergedChild = true;
    }

    private static XlsxCellModel GetOrCreateCell(XlsxSheetModel sheet, int row, int column)
    {
        if (sheet.Cells.TryGetValue((row, column), out var existing)) return existing;
        var cell = new XlsxCellModel { Row = row, Column = column };
        sheet.Cells[(row, column)] = cell;
        return cell;
    }

    private static string FormatValue(string value, string? numberFormat)
    {
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || string.IsNullOrEmpty(numberFormat)) return value;
        if (numberFormat.Contains('%'))
        {
            var decimals = DecimalPlaces(numberFormat);
            return (number * 100).ToString($"F{decimals}", CultureInfo.InvariantCulture) + "%";
        }
        if (numberFormat.Contains("yy", StringComparison.OrdinalIgnoreCase) || numberFormat.Contains("dd", StringComparison.OrdinalIgnoreCase))
        {
            try { return DateTime.FromOADate(number).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture); } catch { return value; }
        }
        if (numberFormat.Contains("#,##", StringComparison.Ordinal)) return number.ToString("N" + DecimalPlaces(numberFormat), CultureInfo.InvariantCulture);
        if (numberFormat.Contains("0", StringComparison.Ordinal)) return number.ToString("F" + DecimalPlaces(numberFormat), CultureInfo.InvariantCulture);
        return value;
    }

    private static int DecimalPlaces(string format)
    {
        var decimalIndex = format.IndexOf('.');
        if (decimalIndex < 0) return 0;
        return format.Skip(decimalIndex + 1).TakeWhile(x => x is '0' or '#').Count();
    }

    private static string? BuiltinNumberFormat(int id) => id switch
    {
        9 => "0%", 10 => "0.00%", 14 => "m/d/yy", 22 => "m/d/yy h:mm", 37 => "#,##0", 38 => "#,##0", 39 => "#,##0.00", 40 => "#,##0.00", _ => null
    };
    private static string ApplyTint(string rgb, double tint)
    {
        if (tint == 0 || rgb.Length != 6) return rgb;
        var r = Convert.ToInt32(rgb[..2], 16); var g = Convert.ToInt32(rgb.Substring(2, 2), 16); var b = Convert.ToInt32(rgb.Substring(4, 2), 16);
        int Shift(int value) => (int)Math.Round(tint < 0 ? value * (1 + tint) : value + (255 - value) * tint);
        return $"{Math.Clamp(Shift(r), 0, 255):X2}{Math.Clamp(Shift(g), 0, 255):X2}{Math.Clamp(Shift(b), 0, 255):X2}";
    }
    private static string IndexedColor(int index) => index switch { 10 => "FF0000", 12 => "0000FF", 23 => "C0C0C0", 22 => "808080", _ => "000000" };
    private static double BorderWidth(XElement? side) => (string?)side?.Attribute("style") switch { null => 0, "medium" or "thick" or "double" => 2, _ => 1 };
    private static TextAlignment ParseHorizontal(string? value) => value switch { "center" or "centerContinuous" => TextAlignment.Center, "right" => TextAlignment.Right, "justify" => TextAlignment.Justify, _ => TextAlignment.Left };
    private static VerticalAlignment ParseVertical(string? value) => value switch { "top" => VerticalAlignment.Top, "bottom" => VerticalAlignment.Bottom, _ => VerticalAlignment.Center };
    private static bool TryParseAddress(string? address, out int column, out int row)
    {
        column = row = 0;
        if (string.IsNullOrWhiteSpace(address)) return false;
        var letters = new string(address.TakeWhile(char.IsLetter).ToArray());
        if (!int.TryParse(address[letters.Length..], out var oneBasedRow) || letters.Length == 0) return false;
        foreach (var item in letters.ToUpperInvariant()) column = column * 26 + item - 'A' + 1;
        column--; row = oneBasedRow - 1; return column >= 0 && row >= 0;
    }
    private static Dictionary<string, string> ReadRelationships(ZipArchive archive, string sourcePath)
    {
        var directory = Path.GetDirectoryName(sourcePath)?.Replace('\\', '/') ?? string.Empty;
        var relationshipsPath = $"{directory}/_rels/{Path.GetFileName(sourcePath)}.rels";
        var entry = archive.GetEntry(relationshipsPath);
        if (entry is null) return [];
        using var stream = entry.Open(); var xml = XDocument.Load(stream);
        return xml.Root?.Elements().Where(x => x.Attribute("Id") is not null && x.Attribute("Target") is not null)
            .ToDictionary(x => (string)x.Attribute("Id")!, x => ResolvePath(sourcePath, (string)x.Attribute("Target")!), StringComparer.Ordinal) ?? [];
    }
    private static XDocument ReadXml(ZipArchive archive, string part) { var entry = archive.GetEntry(part) ?? throw new InvalidDataException($"XLSX part missing: {part}"); using var stream = entry.Open(); return XDocument.Load(stream); }
    private static string ResolvePath(string source, string target) => new Uri(new Uri("http://xlsx/" + source), target).AbsolutePath.TrimStart('/');
    private static FontData AtFont(IReadOnlyList<FontData> values, int index) => index >= 0 && index < values.Count ? values[index] : default;
    private static BorderData AtBorder(IReadOnlyList<BorderData> values, int index) => index >= 0 && index < values.Count ? values[index] : default;
    private static string? AtString(IReadOnlyList<string?> values, int index) => index >= 0 && index < values.Count ? values[index] : null;
    private static XlsxTextRun CloneRun(XlsxTextRun run) => new() { Text = run.Text, Foreground = run.Foreground, FontSize = run.FontSize, FontFamily = run.FontFamily, FontWeight = run.FontWeight, FontStyle = run.FontStyle, Underline = run.Underline };
    private static bool AttributeTrue(XElement? element, string name) => (string?)element?.Attribute(name) is "1" or "true";
    private static int IntAttribute(XElement? element, string name, int fallback) => int.TryParse((string?)element?.Attribute(name), out var value) ? value : fallback;
    private static double DoubleAttribute(XElement? element, string name, double fallback) => double.TryParse((string?)element?.Attribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out var value) ? value : fallback;
    private static long LongAttribute(XElement? element, string name, long fallback) => long.TryParse((string?)element?.Attribute(name), out var value) ? value : fallback;
    private static int IntElement(XElement? parent, XName name, int fallback) => int.TryParse(parent?.Element(name)?.Value, out var value) ? value : fallback;
    private static long LongElement(XElement? parent, XName name, long fallback) => long.TryParse(parent?.Element(name)?.Value, out var value) ? value : fallback;
    private readonly record struct FontData(string? Family, double Size, string? Color, bool Bold, bool Italic, bool Underline);
    private readonly record struct BorderData(Thickness Thickness, string? Color);
}
