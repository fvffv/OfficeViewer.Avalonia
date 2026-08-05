using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Avalonia;
using global::Avalonia.Layout;
using global::Avalonia.Media;

namespace OfficeViewer.Avalonia.Pptx;

/// <summary>
/// Native PresentationML reader. It follows the relationship-driven PPTX model used by
/// pptx-preview: presentation order, per-slide relationship targets, DrawingML shape trees,
/// theme colours, runs, pictures and connectors are read without a web/JavaScript runtime.
/// </summary>
internal sealed class PptxParser
{
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private readonly Dictionary<string, string> _themeColors = new(StringComparer.OrdinalIgnoreCase);

    public PptxDocumentModel Parse(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        return Parse(archive, path, null);
    }

    public PptxDocumentModel Parse(byte[] document)
    {
        ArgumentNullException.ThrowIfNull(document);
        using var archive = new ZipArchive(new MemoryStream(document, writable: false), ZipArchiveMode.Read);
        return Parse(archive, null, document);
    }

    private PptxDocumentModel Parse(ZipArchive archive, string? sourcePath, byte[]? sourceBytes)
    {
        ReadTheme(archive);
        var presentation = ReadXml(archive, "ppt/presentation.xml");
        var size = presentation.Root?.Element(P + "sldSz");
        var result = new PptxDocumentModel
        {
            SourcePath = sourcePath,
            SourceBytes = sourceBytes,
            Width = EmuToPixels(LongAttribute(size, "cx", 12_192_000)),
            Height = EmuToPixels(LongAttribute(size, "cy", 6_858_000))
        };

        var presentationRelationships = ReadRelationships(archive, "ppt/presentation.xml");
        foreach (var slideId in presentation.Descendants(P + "sldId"))
        {
            var relationshipId = (string?)slideId.Attribute(R + "id");
            if (relationshipId is null || !presentationRelationships.TryGetValue(relationshipId, out var slidePath)) continue;
            result.Slides.Add(ParseSlide(archive, slidePath));
        }
        return result;
    }

    private PptxSlideModel ParseSlide(ZipArchive archive, string slidePath)
    {
        var xml = ReadXml(archive, slidePath);
        var relationships = ReadRelationships(archive, slidePath);
        var slide = new PptxSlideModel { Background = ReadFill(xml.Root?.Element(P + "cSld")?.Element(P + "bg")?.Element(P + "bgPr")) };
        var tree = xml.Root?.Element(P + "cSld")?.Element(P + "spTree");
        if (tree is not null) ParseTree(archive, tree, PptxTransform.Identity, relationships, slide.Elements);
        return slide;
    }

    private void ParseTree(ZipArchive archive, XElement tree, PptxTransform transform, Dictionary<string, string> relationships, List<PptxElement> output)
    {
        foreach (var element in tree.Elements())
        {
            if (element.Name == P + "grpSp")
            {
                var xfrm = element.Element(P + "grpSpPr")?.Element(A + "xfrm");
                var off = xfrm?.Element(A + "off");
                var ext = xfrm?.Element(A + "ext");
                var childOff = xfrm?.Element(A + "chOff");
                var childExt = xfrm?.Element(A + "chExt");
                var groupLeft = LongAttribute(off, "x", 0);
                var groupTop = LongAttribute(off, "y", 0);
                var childLeft = LongAttribute(childOff, "x", 0);
                var childTop = LongAttribute(childOff, "y", 0);
                var scaleX = transform.ScaleX * SafeRatio(LongAttribute(ext, "cx", 1), LongAttribute(childExt, "cx", 1));
                var scaleY = transform.ScaleY * SafeRatio(LongAttribute(ext, "cy", 1), LongAttribute(childExt, "cy", 1));
                ParseTree(archive, element, new PptxTransform(transform.X(EmuToPixels(groupLeft)) - EmuToPixels(childLeft) * scaleX, transform.Y(EmuToPixels(groupTop)) - EmuToPixels(childTop) * scaleY, scaleX, scaleY), relationships, output);
            }
            else if (element.Name == P + "sp" && TryReadBounds(element.Element(P + "spPr")?.Element(A + "xfrm"), transform, out var bounds))
                output.Add(ParseShape(element, bounds));
            else if (element.Name == P + "pic" && TryReadBounds(element.Element(P + "spPr")?.Element(A + "xfrm"), transform, out var pictureBounds))
            {
                var relationshipId = (string?)element.Descendants(A + "blip").FirstOrDefault()?.Attribute(R + "embed");
                if (relationshipId is not null && relationships.TryGetValue(relationshipId, out var imagePath))
                    output.Add(new PptxPicture
                    {
                        PackagePath = imagePath,
                        Left = pictureBounds.Left, Top = pictureBounds.Top, Width = pictureBounds.Width, Height = pictureBounds.Height,
                        Rotation = pictureBounds.Rotation, FlipHorizontal = pictureBounds.FlipHorizontal, FlipVertical = pictureBounds.FlipVertical
                    });
            }
            else if (element.Name == P + "cxnSp" && TryReadBounds(element.Element(P + "spPr")?.Element(A + "xfrm"), transform, out var lineBounds))
            {
                var line = element.Element(P + "spPr")?.Element(A + "ln");
                output.Add(new PptxLine
                {
                    Stroke = ReadLineColor(line),
                    StrokeThickness = EmuToPixels(LongAttribute(line, "w", 12_700)),
                    Left = lineBounds.Left, Top = lineBounds.Top, Width = lineBounds.Width, Height = lineBounds.Height,
                    Rotation = lineBounds.Rotation, FlipHorizontal = lineBounds.FlipHorizontal, FlipVertical = lineBounds.FlipVertical
                });
            }
            else if (element.Name == P + "graphicFrame" && TryReadBounds(element.Element(P + "xfrm"), transform, out var chartBounds))
            {
                var relationshipId = (string?)element.Descendants(C + "chart").FirstOrDefault()?.Attribute(R + "id");
                if (relationshipId is not null && relationships.TryGetValue(relationshipId, out var chartPath))
                {
                    var chart = ParseChart(archive, chartPath, chartBounds);
                    if (chart is not null) output.Add(chart);
                }
            }
        }
    }

    private PptxShape ParseShape(XElement shape, PptxBounds bounds)
    {
        var properties = shape.Element(P + "spPr");
        var textBody = shape.Element(P + "txBody");
        var bodyProperties = textBody?.Element(A + "bodyPr");
        var result = new PptxShape
        {
            Geometry = (string?)properties?.Element(A + "prstGeom")?.Attribute("prst") ?? "rect",
            Fill = ReadFill(properties),
            Stroke = ReadLineColor(properties?.Element(A + "ln")),
            StrokeThickness = EmuToPixels(LongAttribute(properties?.Element(A + "ln"), "w", 0)),
            TextInsets = new Thickness(
                EmuToPixels(LongAttribute(bodyProperties, "lIns", 0)),
                EmuToPixels(LongAttribute(bodyProperties, "tIns", 0)),
                EmuToPixels(LongAttribute(bodyProperties, "rIns", 0)),
                EmuToPixels(LongAttribute(bodyProperties, "bIns", 0))),
            VerticalAlignment = ParseVerticalAlignment((string?)bodyProperties?.Attribute("anchor")),
            Left = bounds.Left, Top = bounds.Top, Width = bounds.Width, Height = bounds.Height,
            Rotation = bounds.Rotation, FlipHorizontal = bounds.FlipHorizontal, FlipVertical = bounds.FlipVertical
        };
        if (textBody is not null)
            foreach (var paragraph in textBody.Elements(A + "p")) result.Paragraphs.Add(ParseParagraph(paragraph));
        return result;
    }

    private PptxParagraph ParseParagraph(XElement paragraph)
    {
        var properties = paragraph.Element(A + "pPr");
        var result = new PptxParagraph
        {
            Alignment = ParseTextAlignment((string?)properties?.Attribute("algn")),
            SpaceAfter = EmuToPixels(LongAttribute(properties?.Element(A + "spcAft")?.Element(A + "spcPts"), "val", 0) * 127)
        };
        foreach (var child in paragraph.Elements())
        {
            if (child.Name == A + "br") result.Runs.Add(new PptxTextRun { Text = "\n" });
            if (child.Name != A + "r" && child.Name != A + "fld") continue;
            var propertiesRun = child.Element(A + "rPr");
            var text = child.Element(A + "t")?.Value;
            if (!string.IsNullOrEmpty(text)) result.Runs.Add(new PptxTextRun
            {
                Text = text,
                FontFamily = (string?)propertiesRun?.Attribute("ea") ?? (string?)propertiesRun?.Attribute("latin"),
                FontSize = LongAttribute(propertiesRun, "sz", 0) is var points && points > 0 ? points / 100d * 96d / 72d : null,
                FontWeight = AttributeTrue(propertiesRun, "b") ? FontWeight.Bold : null,
                FontStyle = AttributeTrue(propertiesRun, "i") ? FontStyle.Italic : null,
                Underline = (string?)propertiesRun?.Attribute("u") is { } underline && underline != "none",
                Foreground = ReadFill(propertiesRun)
            });
        }
        return result;
    }

    private bool TryReadBounds(XElement? xfrm, PptxTransform transform, out PptxBounds bounds)
    {
        bounds = default;
        var off = xfrm?.Element(A + "off");
        var ext = xfrm?.Element(A + "ext");
        if (off is null || ext is null) return false;
        var left = LongAttribute(off, "x", 0);
        var top = LongAttribute(off, "y", 0);
        bounds = new PptxBounds(
            transform.X(EmuToPixels(left)), transform.Y(EmuToPixels(top)),
            EmuToPixels(LongAttribute(ext, "cx", 0)) * transform.ScaleX,
            EmuToPixels(LongAttribute(ext, "cy", 0)) * transform.ScaleY,
            LongAttribute(xfrm, "rot", 0) / 60_000d, AttributeTrue(xfrm, "flipH"), AttributeTrue(xfrm, "flipV"));
        return bounds.Width > 0 && bounds.Height > 0;
    }

    private void ReadTheme(ZipArchive archive)
    {
        var entry = archive.GetEntry("ppt/theme/theme1.xml");
        if (entry is null) return;
        using var stream = entry.Open();
        var theme = XDocument.Load(stream);
        var scheme = theme.Descendants(A + "clrScheme").FirstOrDefault();
        if (scheme is null) return;
        foreach (var item in scheme.Elements())
        {
            var color = item.Descendants(A + "srgbClr").Attributes("val").FirstOrDefault()?.Value
                ?? item.Descendants(A + "sysClr").Attributes("lastClr").FirstOrDefault()?.Value;
            if (color is not null) _themeColors[item.Name.LocalName] = color;
        }

        // DrawingML uses the semantic names bg1/tx1 in shapes, while the
        // colour scheme itself exposes their values as lt1/dk1.  Treat them as
        // aliases here so a background overlay is not silently transparent.
        CopyThemeAlias("bg1", "lt1");
        CopyThemeAlias("tx1", "dk1");
        CopyThemeAlias("bg2", "lt2");
        CopyThemeAlias("tx2", "dk2");
    }

    private PptxChart? ParseChart(ZipArchive archive, string path, PptxBounds bounds)
    {
        var chartXml = ReadXml(archive, path);
        var series = chartXml.Descendants(C + "barChart").Elements(C + "ser").FirstOrDefault();
        if (series is null) return null;
        var result = new PptxChart
        {
            Left = bounds.Left, Top = bounds.Top, Width = bounds.Width, Height = bounds.Height,
            Rotation = bounds.Rotation, FlipHorizontal = bounds.FlipHorizontal, FlipVertical = bounds.FlipVertical,
            BarFill = ReadFill(series.Element(C + "spPr")),
            Title = string.Concat(chartXml.Descendants(C + "title").Descendants(A + "t").Select(x => x.Value))
        };
        foreach (var point in series.Descendants(C + "cat").Descendants(C + "pt").OrderBy(x => LongAttribute(x, "idx", 0)))
            result.Categories.Add(point.Element(C + "v")?.Value ?? string.Empty);
        foreach (var point in series.Descendants(C + "val").Descendants(C + "pt").OrderBy(x => LongAttribute(x, "idx", 0)))
            if (double.TryParse(point.Element(C + "v")?.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var value))
                result.Values.Add(value);
        return result.Values.Count == 0 ? null : result;
    }

    private PptxFill? ReadFill(XElement? properties)
    {
        var solidFill = properties?.Elements().FirstOrDefault(x => x.Name == A + "solidFill");
        if (solidFill is not null)
        {
            var color = ReadColor(solidFill.Elements().FirstOrDefault());
            return color is null ? null : new PptxFill { Color = color };
        }

        var gradient = properties?.Elements().FirstOrDefault(x => x.Name == A + "gradFill");
        if (gradient is null) return null;
        var result = new PptxFill { GradientAngle = LongAttribute(gradient.Element(A + "lin"), "ang", 0) / 60_000d };
        foreach (var stop in gradient.Element(A + "gsLst")?.Elements(A + "gs") ?? [])
        {
            var color = ReadColor(stop.Elements().FirstOrDefault());
            if (color is not null)
                result.GradientStops.Add(new PptxGradientStop(LongAttribute(stop, "pos", 0) / 100_000d, color));
        }
        return result.GradientStops.Count == 0 ? null : result;
    }

    private string? ReadColor(XElement? colorElement)
    {
        if (colorElement is null) return null;
        var rgb = colorElement.Name == A + "srgbClr"
            ? (string?)colorElement.Attribute("val")
            : colorElement.Name == A + "sysClr"
                ? (string?)colorElement.Attribute("lastClr")
                : colorElement.Name == A + "prstClr"
                    ? PresetColor((string?)colorElement.Attribute("val"))
                    : (string?)colorElement.Attribute("val") is { } scheme && _themeColors.TryGetValue(scheme, out var themeColor)
                        ? themeColor
                        : null;
        if (string.IsNullOrWhiteSpace(rgb)) return null;
        var alphaValue = LongAttribute(colorElement.Element(A + "alpha"), "val", 100_000);
        var alpha = (byte)Math.Clamp((int)Math.Round(alphaValue * 255d / 100_000d), 0, 255);
        return alpha == 255 ? rgb : $"{alpha:X2}{rgb}";
    }

    private static string? PresetColor(string? value) => value switch
    {
        "black" => "000000", "white" => "FFFFFF", "gray" => "808080", "dkGray" => "404040", "ltGray" => "C0C0C0",
        _ => null
    };

    private void CopyThemeAlias(string alias, string source)
    {
        if (!_themeColors.ContainsKey(alias) && _themeColors.TryGetValue(source, out var color))
            _themeColors[alias] = color;
    }

    private PptxFill? ReadLineColor(XElement? line) => ReadFill(line);

    private static Dictionary<string, string> ReadRelationships(ZipArchive archive, string partPath)
    {
        var relationshipsPath = Path.ChangeExtension(partPath, ".xml.rels").Replace("/slides/", "/slides/_rels/").Replace("/slideLayouts/", "/slideLayouts/_rels/").Replace("/presentation.xml.rels", "/_rels/presentation.xml.rels");
        var entry = archive.GetEntry(relationshipsPath);
        if (entry is null) return new(StringComparer.Ordinal);
        using var stream = entry.Open();
        var xml = XDocument.Load(stream);
        return xml.Root?.Elements().Where(x => (string?)x.Attribute("TargetMode") != "External")
            .Where(x => x.Attribute("Id") is not null && x.Attribute("Target") is not null)
            .ToDictionary(x => (string)x.Attribute("Id")!, x => ResolvePartPath(partPath, (string)x.Attribute("Target")!), StringComparer.Ordinal)
            ?? new(StringComparer.Ordinal);
    }

    private static XDocument ReadXml(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path) ?? throw new InvalidDataException($"PPTX part missing: {path}");
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static string ResolvePartPath(string sourcePart, string target)
    {
        var sourceUri = new Uri("http://pptx/" + sourcePart);
        return new Uri(sourceUri, target).AbsolutePath.TrimStart('/');
    }

    private static TextAlignment? ParseTextAlignment(string? value) => value switch
    {
        "ctr" => TextAlignment.Center, "r" => TextAlignment.Right, "just" => TextAlignment.Justify, "l" => TextAlignment.Left, _ => null
    };
    private static VerticalAlignment ParseVerticalAlignment(string? value) => value switch
    {
        "ctr" => VerticalAlignment.Center, "b" => VerticalAlignment.Bottom, _ => VerticalAlignment.Top
    };
    private static bool AttributeTrue(XElement? element, string name) => (string?)element?.Attribute(name) is "1" or "true";
    private static long LongAttribute(XElement? element, string name, long fallback) => long.TryParse((string?)element?.Attribute(name), out var value) ? value : fallback;
    private static double EmuToPixels(long emu) => emu * 96d / 914400d;
    private static double SafeRatio(long numerator, long denominator) => denominator == 0 ? 1 : numerator / (double)denominator;
}
