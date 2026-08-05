using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using global::Avalonia.Media;

namespace OfficeViewer.Avalonia.Pdf;

/// <summary>
/// A small, dependency-free PDF reader for the common document subset used by
/// office exports: classic objects/xref, page trees, Flate streams, Type0
/// ToUnicode fonts, standard text operators, straight paths and RGB images.
/// It deliberately mirrors PDF.js' parse-to-display-list split, while leaving
/// native rendering to PdfViewer/Avalonia controls.
/// </summary>
internal sealed class PdfDocumentModel : IDisposable
{
    private PdfParser? _parser;
    private PdfSource? _source;
    private PdfPageInfo[]? _pageInfos;
    private PdfPageModel?[]? _pages;

    internal PdfDocumentModel(PdfParser parser, PdfSource source, PdfPageInfo[] pageInfos)
    {
        _parser = parser;
        _source = source;
        _pageInfos = pageInfos;
        _pages = new PdfPageModel?[pageInfos.Length];
    }

    public int PageCount => _pageInfos?.Length ?? 0;

    public PdfPageInfo GetPageInfo(int index) => _pageInfos?[index] ?? throw new ObjectDisposedException(nameof(PdfDocumentModel));

    public PdfPageModel GetPage(int index)
    {
        var pages = _pages ?? throw new ObjectDisposedException(nameof(PdfDocumentModel));
        if (pages[index] is { } page) return page;
        var parser = _parser ?? throw new ObjectDisposedException(nameof(PdfDocumentModel));
        var source = _source ?? throw new ObjectDisposedException(nameof(PdfDocumentModel));
        return pages[index] = parser.ParsePage(source, GetPageInfo(index).Reference, index + 1);
    }

    // Clearing a viewer drops source bytes, decoded commands and image payloads.
    public void Dispose()
    {
        _pages = null;
        _pageInfos = null;
        _source = null;
        _parser = null;
    }
}

internal sealed record PdfPageInfo(int Number, PdfReference Reference, double Width, double Height);

internal sealed record PdfPageModel(
    int Number,
    double Width,
    double Height,
    IReadOnlyList<PdfTextSegment> Text,
    IReadOnlyList<PdfLineSegment> Lines,
    IReadOnlyList<PdfFilledPolygon> Fills,
    IReadOnlyList<PdfImageModel> Images,
    IReadOnlyList<PdfPageDrawOperation> DrawOperations);

internal abstract record PdfPageDrawOperation;
internal sealed record PdfPageTextDraw(PdfTextSegment Text) : PdfPageDrawOperation;
internal sealed record PdfPageLineDraw(PdfLineSegment Line) : PdfPageDrawOperation;
internal sealed record PdfPageFillDraw(PdfFilledPolygon Fill) : PdfPageDrawOperation;
internal sealed record PdfPageImageDraw(PdfImageModel Image) : PdfPageDrawOperation;

internal sealed record PdfTextSegment(
    string Text,
    double Left,
    double Baseline,
    double FontSize,
    string? FontFamily,
    Color Color);

internal sealed record PdfLineSegment(
    double StartX,
    double StartY,
    double EndX,
    double EndY,
    double Thickness,
    Color Color);

internal sealed record PdfFilledPolygon(IReadOnlyList<PdfPoint> Points, Color Color);

internal sealed record PdfImagePayload(
    byte[] EncodedBytes,
    IReadOnlyList<string> Filters,
    int PixelWidth,
    int PixelHeight,
    int BitsPerComponent,
    string ColorSpace,
    byte[]? SoftMaskBytes,
    IReadOnlyList<string>? SoftMaskFilters,
    byte[]? ColorKeyMask);

internal sealed record PdfImageModel(PdfImagePayload Payload, double Left, double Top, double Width, double Height, double Opacity, bool IsWatermark);

/// <summary>
/// PDF resources are scoped. A Form XObject can have a different font/image table
/// from its page, so resolving only the page-level resource dictionary drops the
/// complete contents of many exported PDFs.
/// </summary>
internal sealed class PdfResourceSet
{
    private readonly PdfResourceSet? _parent;
    public PdfResourceSet(PdfResourceSet? parent = null) => _parent = parent;
    public Dictionary<string, PdfFont> Fonts { get; } = new(StringComparer.Ordinal);
    public Dictionary<string, PdfReference> Images { get; } = new(StringComparer.Ordinal);
    public Dictionary<string, PdfFormXObject> Forms { get; } = new(StringComparer.Ordinal);
    public Dictionary<string, PdfExtGraphicsState> GraphicsStates { get; } = new(StringComparer.Ordinal);

    public bool TryGetFont(string key, out PdfFont? font) =>
        Fonts.TryGetValue(key, out font) || (_parent?.TryGetFont(key, out font) ?? false);

    public bool TryGetImage(string key, out PdfReference? image) =>
        Images.TryGetValue(key, out image) || (_parent?.TryGetImage(key, out image) ?? false);

    public bool TryGetForm(string key, out PdfFormXObject? form) =>
        Forms.TryGetValue(key, out form) || (_parent?.TryGetForm(key, out form) ?? false);

    public bool TryGetGraphicsState(string key, out PdfExtGraphicsState? graphicsState) =>
        GraphicsStates.TryGetValue(key, out graphicsState) || (_parent?.TryGetGraphicsState(key, out graphicsState) ?? false);
}

internal sealed record PdfFormXObject(byte[] Content, PdfResourceSet Resources, PdfMatrix Matrix);
internal sealed record PdfExtGraphicsState(double StrokeOpacity, double FillOpacity);

internal sealed class PdfParser
{
    // Image streams are frequently shared by hundreds of pages. Keeping one
    // payload per indirect object prevents a multi-page PDF from retaining a
    // duplicate JPEG/Flate byte array for every use.
    private readonly Dictionary<int, PdfImagePayload> _imagePayloadCache = [];
    private readonly Dictionary<int, PdfFont> _fontCache = [];

    public PdfDocumentModel Parse(string path)
    {
        return Parse(PdfSource.Read(path));
    }

    public PdfDocumentModel Parse(byte[] document)
    {
        ArgumentNullException.ThrowIfNull(document);
        return Parse(PdfSource.Read(document));
    }

    private PdfDocumentModel Parse(PdfSource source)
    {
        var catalog = source.Objects.Values
            .Select(item => item.Dictionary)
            .OfType<PdfDictionary>()
            .FirstOrDefault(dictionary => string.Equals(source.NameOf(dictionary.Get("Type")), "Catalog", StringComparison.Ordinal));
        if (catalog is null) throw new InvalidDataException("PDF catalog was not found.");
        if (catalog.Get("Pages") is not PdfReference pagesReference) throw new InvalidDataException("PDF page tree was not found.");

        var pageReferences = new List<PdfReference>();
        CollectPageReferences(source, pagesReference, pageReferences);
        var pageInfos = new PdfPageInfo[pageReferences.Count];
        for (var index = 0; index < pageReferences.Count; index++)
            pageInfos[index] = ReadPageInfo(source, pageReferences[index], index + 1);
        return new PdfDocumentModel(this, source, pageInfos);
    }

    private static void CollectPageReferences(PdfSource source, PdfReference nodeReference, ICollection<PdfReference> destination)
    {
        var node = source.DictionaryOf(nodeReference);
        var type = source.NameOf(node.Get("Type"));
        if (string.Equals(type, "Page", StringComparison.Ordinal))
        {
            destination.Add(nodeReference);
            return;
        }
        if (node.Get("Kids") is not PdfArray kids) return;
        foreach (var kid in kids.Values.OfType<PdfReference>()) CollectPageReferences(source, kid, destination);
    }

    internal PdfPageModel ParsePage(PdfSource source, PdfReference pageReference, int number)
    {
        var page = source.DictionaryOf(pageReference);
        var pageInfo = ReadPageInfo(source, pageReference, number);
        var width = pageInfo.Width;
        var height = pageInfo.Height;
        var resources = ReadResources(source, source.DictionaryOf(InheritedValue(source, pageReference, "Resources")), null, []);
        var content = ReadPageContent(source, page.Get("Contents"));
        var commands = new PdfContentInterpreter(this, source, resources).Interpret(content);

        var pageText = commands.Text;
        var pageLines = commands.Lines;
        var pageFills = commands.Fills.Select(fill => ToPageFill(fill, height)).ToArray();
        var pageImages = commands.Images.Select(image => ToPageImage(image, height)).ToArray();
        var pageDrawOperations = commands.DrawOperations.Select<PdfDrawOperation, PdfPageDrawOperation>(operation => operation switch
        {
            PdfTextDrawOperation text => new PdfPageTextDraw(text.Text),
            PdfLineDrawOperation line => new PdfPageLineDraw(line.Line),
            PdfFillDrawOperation fill => new PdfPageFillDraw(ToPageFill(fill.Fill, height)),
            PdfImageDrawOperation image => new PdfPageImageDraw(ToPageImage(image.Image, height)),
            _ => throw new InvalidDataException("PDF display operation is not supported.")
        }).ToArray();

        return new PdfPageModel(
            number,
            width,
            height,
            pageText,
            pageLines,
            pageFills,
            pageImages,
            pageDrawOperations);
    }

    private static PdfFilledPolygon ToPageFill(PdfFilledPolygon fill, double pageHeight) =>
        new(fill.Points.Select(point => new PdfPoint(point.X, pageHeight - point.Y)).ToArray(), fill.Color);

    private static PdfPageInfo ReadPageInfo(PdfSource source, PdfReference pageReference, int number)
    {
        var mediaBox = source.ArrayOf(InheritedValue(source, pageReference, "MediaBox"));
        if (mediaBox is null || mediaBox.Values.Count < 4) throw new InvalidDataException("PDF page MediaBox is missing.");
        var left = source.NumberOf(mediaBox.Values[0]);
        var bottom = source.NumberOf(mediaBox.Values[1]);
        var width = source.NumberOf(mediaBox.Values[2]) - left;
        var height = source.NumberOf(mediaBox.Values[3]) - bottom;
        if (width <= 0 || height <= 0) throw new InvalidDataException("PDF page MediaBox is invalid.");
        return new PdfPageInfo(number, pageReference, width, height);
    }

    private static PdfImageModel ToPageImage(PdfPlacedImage image, double pageHeight)
    {
        var p0 = image.Transform.Transform(0, 0);
        var p1 = image.Transform.Transform(1, 0);
        var p2 = image.Transform.Transform(0, 1);
        var p3 = image.Transform.Transform(1, 1);
        var minX = Math.Min(Math.Min(p0.X, p1.X), Math.Min(p2.X, p3.X));
        var maxX = Math.Max(Math.Max(p0.X, p1.X), Math.Max(p2.X, p3.X));
        var minY = Math.Min(Math.Min(p0.Y, p1.Y), Math.Min(p2.Y, p3.Y));
        var maxY = Math.Max(Math.Max(p0.Y, p1.Y), Math.Max(p2.Y, p3.Y));
        return new PdfImageModel(image.Payload, minX, pageHeight - maxY, maxX - minX, maxY - minY, image.Opacity, image.IsWatermark);
    }

    private static PdfValue? InheritedValue(PdfSource source, PdfReference pageReference, string key)
    {
        PdfReference? current = pageReference;
        var guard = 0;
        while (current is not null && guard++ < 64)
        {
            var dictionary = source.DictionaryOf(current);
            if (dictionary.Get(key) is { } value) return value;
            current = dictionary.Get("Parent") as PdfReference;
        }
        return null;
    }

    private PdfResourceSet ReadResources(PdfSource source, PdfDictionary? resources, PdfResourceSet? parent, HashSet<int> formStack)
    {
        var result = new PdfResourceSet(parent);

        var fontDictionary = source.DictionaryOf(resources?.Get("Font"));
        if (fontDictionary is not null)
        {
            foreach (var pair in fontDictionary.Values)
            {
                var font = ReadFont(source, pair.Value);
                if (font is not null) result.Fonts[pair.Key] = font;
            }
        }

        var graphicsStateDictionary = source.DictionaryOf(resources?.Get("ExtGState"));
        if (graphicsStateDictionary is not null)
        {
            foreach (var pair in graphicsStateDictionary.Values)
            {
                var dictionary = source.DictionaryOf(pair.Value);
                if (dictionary is null) continue;
                result.GraphicsStates[pair.Key] = new PdfExtGraphicsState(
                    Math.Clamp(source.NumberOf(dictionary.Get("CA"), 1), 0, 1),
                    Math.Clamp(source.NumberOf(dictionary.Get("ca"), 1), 0, 1));
            }
        }

        var xObjects = source.DictionaryOf(resources?.Get("XObject"));
        if (xObjects is null) return result;
        foreach (var pair in xObjects.Values)
        {
            if (pair.Value is not PdfReference reference) continue;
            var item = source.ObjectOf(reference);
            var dictionary = item?.Dictionary;
            if (item is null || dictionary is null) continue;
            var subtype = source.NameOf(dictionary.Get("Subtype"));
            if (string.Equals(subtype, "Image", StringComparison.Ordinal))
            {
                result.Images[pair.Key] = reference;
                continue;
            }
            if (!string.Equals(subtype, "Form", StringComparison.Ordinal) || !formStack.Add(reference.Number)) continue;
            try
            {
                var formResources = ReadResources(source, source.DictionaryOf(dictionary.Get("Resources")), result, formStack);
                result.Forms[pair.Key] = new PdfFormXObject(
                    source.DecodedStreamOf(item),
                    formResources,
                    MatrixOf(source.ArrayOf(dictionary.Get("Matrix"))));
            }
            finally
            {
                formStack.Remove(reference.Number);
            }
        }
        return result;
    }

    private PdfImagePayload ReadImagePayload(PdfSource source, PdfReference reference, PdfRawObject item, PdfDictionary dictionary)
    {
        if (_imagePayloadCache.TryGetValue(reference.Number, out var cached)) return cached;
        var pixelWidth = (int)Math.Max(1, source.NumberOf(dictionary.Get("Width")));
        var pixelHeight = (int)Math.Max(1, source.NumberOf(dictionary.Get("Height")));
        var bits = (int)Math.Max(1, source.NumberOf(dictionary.Get("BitsPerComponent"), 8));
        var colorSpace = source.NameOf(dictionary.Get("ColorSpace")) ?? "DeviceRGB";
        var softMask = dictionary.Get("SMask") as PdfReference;
        var maskObject = softMask is null ? null : source.ObjectOf(softMask);
        var payload = new PdfImagePayload(
            source.EncodedStreamOf(item),
            source.FilterNames(item),
            pixelWidth,
            pixelHeight,
            bits,
            colorSpace,
            maskObject is null ? null : source.EncodedStreamOf(maskObject),
            maskObject is null ? null : source.FilterNames(maskObject),
            ToColorKeyMask(source.ArrayOf(dictionary.Get("Mask"))));
        _imagePayloadCache[reference.Number] = payload;
        return payload;
    }

    internal PdfImagePayload GetImagePayload(PdfSource source, PdfReference reference)
    {
        if (_imagePayloadCache.TryGetValue(reference.Number, out var cached)) return cached;
        var item = source.ObjectOf(reference) ?? throw new InvalidDataException($"PDF image object {reference.Number} is missing.");
        var dictionary = item.Dictionary ?? throw new InvalidDataException($"PDF image object {reference.Number} has no dictionary.");
        return ReadImagePayload(source, reference, item, dictionary);
    }

    private PdfFont? ReadFont(PdfSource source, PdfValue reference)
    {
        if (reference is PdfReference fontReference && _fontCache.TryGetValue(fontReference.Number, out var cached)) return cached;
        var dictionary = source.DictionaryOf(reference);
        if (dictionary is null) return null;
        var isComposite = string.Equals(source.NameOf(dictionary.Get("Subtype")), "Type0", StringComparison.Ordinal);
        var widths = isComposite
            ? ReadCidWidths(source, dictionary, out var defaultWidth)
            : ReadSimpleWidths(source, dictionary, out defaultWidth);
        var font = new PdfFont(
            NormalizeFontName(source.NameOf(dictionary.Get("BaseFont"))),
            ReadToUnicodeMap(source, dictionary.Get("ToUnicode")),
            widths,
            defaultWidth,
            isComposite ? 2 : 1);
        if (reference is PdfReference cachedReference) _fontCache[cachedReference.Number] = font;
        return font;
    }

    private static Dictionary<int, double> ReadSimpleWidths(PdfSource source, PdfDictionary dictionary, out double defaultWidth)
    {
        defaultWidth = 500;
        var firstChar = (int)source.NumberOf(dictionary.Get("FirstChar"));
        var values = source.ArrayOf(dictionary.Get("Widths"));
        var result = new Dictionary<int, double>();
        if (values is null) return result;
        for (var index = 0; index < values.Values.Count; index++)
            if (values.Values[index] is PdfNumber width) result[firstChar + index] = (double)width.Value;
        return result;
    }

    private static Dictionary<int, double> ReadCidWidths(PdfSource source, PdfDictionary dictionary, out double defaultWidth)
    {
        defaultWidth = 1000;
        var descendants = source.ArrayOf(dictionary.Get("DescendantFonts"));
        var descendant = descendants?.Values.Count > 0 ? source.DictionaryOf(descendants.Values[0]) : null;
        if (descendant is null) return [];
        defaultWidth = source.NumberOf(descendant.Get("DW"), 1000);
        var values = source.ArrayOf(descendant.Get("W"));
        var result = new Dictionary<int, double>();
        if (values is null) return result;
        for (var index = 0; index + 1 < values.Values.Count;)
        {
            if (values.Values[index] is not PdfNumber first) { index++; continue; }
            var firstCode = (int)first.Value;
            var following = values.Values[index + 1];
            if (following is PdfArray explicitWidths)
            {
                for (var widthIndex = 0; widthIndex < explicitWidths.Values.Count; widthIndex++)
                    if (explicitWidths.Values[widthIndex] is PdfNumber width) result[firstCode + widthIndex] = (double)width.Value;
                index += 2;
                continue;
            }
            if (following is PdfNumber last && index + 2 < values.Values.Count && values.Values[index + 2] is PdfNumber sharedWidth)
            {
                for (var code = firstCode; code <= (int)last.Value; code++) result[code] = (double)sharedWidth.Value;
                index += 3;
                continue;
            }
            index += 2;
        }
        return result;
    }

    private static PdfMatrix MatrixOf(PdfArray? values)
    {
        if (values is null || values.Values.Count < 6) return PdfMatrix.Identity;
        return new PdfMatrix(
            values.Values[0] is PdfNumber a ? (double)a.Value : 1,
            values.Values[1] is PdfNumber b ? (double)b.Value : 0,
            values.Values[2] is PdfNumber c ? (double)c.Value : 0,
            values.Values[3] is PdfNumber d ? (double)d.Value : 1,
            values.Values[4] is PdfNumber e ? (double)e.Value : 0,
            values.Values[5] is PdfNumber f ? (double)f.Value : 0);
    }

    private static byte[]? ToColorKeyMask(PdfArray? mask)
    {
        if (mask is null || mask.Values.Count < 6) return null;
        var bytes = new byte[6];
        for (var index = 0; index < bytes.Length; index++)
            if (mask.Values[index] is PdfNumber number) bytes[index] = (byte)Math.Clamp((int)number.Value, 0, 255);
            else return null;
        return bytes;
    }

    private static byte[] ReadPageContent(PdfSource source, PdfValue? contents)
    {
        var streams = new List<byte[]>();
        void Append(PdfValue? value)
        {
            if (value is PdfArray array) { foreach (var part in array.Values) Append(part); return; }
            var streamObject = source.ObjectOf(value);
            if (streamObject is not null) streams.Add(source.DecodedStreamOf(streamObject));
        }
        Append(contents);
        if (streams.Count == 0) return [];
        if (streams.Count == 1) return streams[0];
        var length = streams.Sum(item => item.Length + 1);
        var merged = new byte[length];
        var offset = 0;
        foreach (var stream in streams)
        {
            Buffer.BlockCopy(stream, 0, merged, offset, stream.Length);
            offset += stream.Length;
            merged[offset++] = (byte)'\n';
        }
        return merged;
    }

    private static Dictionary<string, string> ReadToUnicodeMap(PdfSource source, PdfValue? reference)
    {
        var item = source.ObjectOf(reference);
        if (item is null) return [];
        var text = Encoding.Latin1.GetString(source.DecodedStreamOf(item));
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var lineStart = 0;
        var section = CMapSection.None;
        while (lineStart < text.Length)
        {
            var lineEnd = text.IndexOf('\n', lineStart);
            if (lineEnd < 0) lineEnd = text.Length;
            var line = text[lineStart..lineEnd];
            if (line.Contains("beginbfchar", StringComparison.Ordinal))
            {
                var searchStart = 0;
                var hasOpenSection = false;
                while (true)
                {
                    var begin = line.IndexOf("beginbfchar", searchStart, StringComparison.Ordinal);
                    if (begin < 0) break;
                    var contentStart = begin + "beginbfchar".Length;
                    var contentEnd = line.IndexOf("endbfchar", contentStart, StringComparison.Ordinal);
                    var mappings = contentEnd >= 0 ? line[contentStart..contentEnd] : line[contentStart..];
                    AddBfCharEntries(map, mappings);
                    if (contentEnd < 0) { hasOpenSection = true; break; }
                    searchStart = contentEnd + "endbfchar".Length;
                }
                section = hasOpenSection ? CMapSection.Char : CMapSection.None;
            }
            else if (line.Contains("beginbfrange", StringComparison.Ordinal)) section = CMapSection.Range;
            else if (line.Contains("endbfchar", StringComparison.Ordinal) || line.Contains("endbfrange", StringComparison.Ordinal)) section = CMapSection.None;
            else if (section == CMapSection.Char)
            {
                AddBfCharEntries(map, line);
            }
            else if (section == CMapSection.Range) AddBfRange(map, line);
            lineStart = lineEnd + 1;
        }
        return map;
    }

    private static void AddBfCharEntries(Dictionary<string, string> map, string text)
    {
        var groups = ExtractHexGroups(text);
        for (var index = 0; index + 1 < groups.Count; index += 2)
            map[groups[index]] = DecodeUnicodeHex(groups[index + 1]);
    }

    private static void AddBfRange(Dictionary<string, string> map, string line)
    {
        var groups = ExtractHexGroups(line);
        if (groups.Count < 3 || !ulong.TryParse(groups[0], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var first) ||
            !ulong.TryParse(groups[1], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var last) || last < first) return;
        // CMaps can legally describe very large ranges. The documents we support
        // use character maps; avoiding an unbounded allocation is important for
        // hostile/corrupt input while still covering the full BMP range.
        var count = Math.Min(last - first + 1, 65_536UL);
        var keyWidth = groups[0].Length;
        var destinations = line.Contains("[", StringComparison.Ordinal) ? groups.Skip(2).ToArray() : null;
        for (ulong index = 0; index < count; index++)
        {
            if (destinations is not null && index >= (ulong)destinations.Length) break;
            var destination = destinations is null ? OffsetUnicode(groups[2], index) : DecodeUnicodeHex(destinations[index]);
            map[(first + index).ToString($"X{keyWidth}", CultureInfo.InvariantCulture)] = destination;
        }
    }

    private static string OffsetUnicode(string destination, ulong offset)
    {
        var text = DecodeUnicodeHex(destination);
        if (string.IsNullOrEmpty(text) || offset == 0) return text;
        var codePoint = char.ConvertToUtf32(text, 0);
        var shifted = codePoint + (long)offset;
        return shifted is > 0 and <= 0x10FFFF ? char.ConvertFromUtf32((int)shifted) : text;
    }

    private enum CMapSection { None, Char, Range }

    private static List<string> ExtractHexGroups(string text)
    {
        var result = new List<string>();
        var position = 0;
        while (position < text.Length)
        {
            var start = text.IndexOf('<', position);
            if (start < 0) break;
            var end = text.IndexOf('>', start + 1);
            if (end < 0) break;
            result.Add(text[(start + 1)..end].Replace(" ", string.Empty, StringComparison.Ordinal));
            position = end + 1;
        }
        return result;
    }

    private static string DecodeUnicodeHex(string text)
    {
        var bytes = PdfSyntax.ToBytes(text);
        if (bytes.Length == 0) return string.Empty;
        return bytes.Length % 2 == 0 ? Encoding.BigEndianUnicode.GetString(bytes) : Encoding.Latin1.GetString(bytes);
    }

    private static IReadOnlyList<PdfTextSegment> GroupText(IReadOnlyList<PdfTextSegment> source)
    {
        var result = new List<PdfTextSegment>();
        foreach (var line in source.GroupBy(item => Math.Round(item.Baseline * 2) / 2).OrderByDescending(group => group.Key))
        {
            TextSegmentBuilder? current = null;
            foreach (var item in line.OrderBy(item => item.Left))
            {
                var canAppend = current is not null
                    && string.Equals(current.FontFamily, item.FontFamily, StringComparison.OrdinalIgnoreCase)
                    && current.Color == item.Color
                    && Math.Abs(current.FontSize - item.FontSize) < .15
                    && item.Left - current.Right <= Math.Max(2, item.FontSize * .8)
                    && item.Left - current.Right >= -1;
                if (!canAppend)
                {
                    current?.AddTo(result);
                    current = new TextSegmentBuilder(item);
                }
                else current!.Append(item.Text);
            }
            current?.AddTo(result);
        }
        return result;
    }

    private static string? NormalizeFontName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return null;
        var separator = name.IndexOf('+');
        return separator >= 0 && separator < name.Length - 1 ? name[(separator + 1)..] : name;
    }

    private sealed class TextSegmentBuilder
    {
        private readonly StringBuilder _text = new();
        public double Left { get; }
        public double Baseline { get; }
        public double FontSize { get; }
        public string? FontFamily { get; }
        public Color Color { get; }
        public double Right { get; private set; }

        public TextSegmentBuilder(PdfTextSegment initial)
        {
            Left = initial.Left;
            Baseline = initial.Baseline;
            FontSize = initial.FontSize;
            FontFamily = initial.FontFamily;
            Color = initial.Color;
            Append(initial.Text);
        }

        public void Append(string text)
        {
            _text.Append(text);
            Right = Left + PdfFont.EstimateAdvance(_text.ToString(), FontSize);
        }

        public void AddTo(ICollection<PdfTextSegment> output)
        {
            if (_text.Length > 0) output.Add(new PdfTextSegment(_text.ToString(), Left, Baseline, FontSize, FontFamily, Color));
        }
    }
}

internal sealed class PdfContentInterpreter(PdfParser parser, PdfSource source, PdfResourceSet resources)
{
    public PdfPageCommands Interpret(byte[] content)
    {
        var output = new PdfPageCommands();
        InterpretInto(content, PdfGraphicsState.Default, output, 0, false);
        return output;
    }

    private void InterpretInto(byte[] content, PdfGraphicsState initialState, PdfPageCommands output, int depth, bool inheritedWatermark)
    {
        if (depth > 24) return; // malformed PDFs can form a cyclic XObject graph.
        var reader = new PdfContentReader(content);
        var operands = new List<PdfContentValue>();
        var state = initialState;
        var saved = new Stack<PdfGraphicsState>();
        var path = new List<PdfLineSegment>();
        var fillPath = new List<PdfFilledPolygon>();
        PdfPoint? currentPoint = null;
        PdfPoint? subpathStart = null;
        var textMatrix = PdfMatrix.Identity;
        var textLineMatrix = PdfMatrix.Identity;
        var fontKey = string.Empty;
        var fontSize = 12d;
        var leading = 0d;
        var characterSpacing = 0d;
        var wordSpacing = 0d;
        var horizontalScaling = 1d;
        var markedContent = new Stack<bool>();
        var isWatermark = inheritedWatermark;

        while (reader.TryRead(out var token))
        {
            if (token.Kind != PdfContentKind.Word) { operands.Add(token); continue; }
            switch (token.Text)
            {
                case "q": saved.Push(state); break;
                case "Q": if (saved.Count > 0) state = saved.Pop(); break;
                case "BMC":
                    markedContent.Push(isWatermark);
                    break;
                case "BDC":
                    markedContent.Push(isWatermark);
                    if (operands.LastOrDefault().Kind == PdfContentKind.Dictionary &&
                        operands[^1].Text.Contains("/Subtype/Watermark", StringComparison.Ordinal))
                        isWatermark = true;
                    break;
                case "EMC":
                    if (markedContent.Count > 0) isWatermark = markedContent.Pop();
                    break;
                case "cm":
                    if (Numbers(operands, 6, out var cm)) state = state with { Transform = state.Transform.Multiply(new PdfMatrix(cm[0], cm[1], cm[2], cm[3], cm[4], cm[5])) };
                    break;
                case "w": if (Number(operands, 1, out var width)) state = state with { LineWidth = Math.Max(.25, width) }; break;
                case "RG": if (Numbers(operands, 3, out var stroke)) state = state with { Stroke = WithOpacity(Rgb(stroke[0], stroke[1], stroke[2]), state.StrokeOpacity) }; break;
                case "rg": if (Numbers(operands, 3, out var fill)) state = state with { Fill = WithOpacity(Rgb(fill[0], fill[1], fill[2]), state.FillOpacity) }; break;
                case "G": if (Number(operands, 1, out var grayStroke)) state = state with { Stroke = WithOpacity(Gray(grayStroke), state.StrokeOpacity) }; break;
                case "g": if (Number(operands, 1, out var grayFill)) state = state with { Fill = WithOpacity(Gray(grayFill), state.FillOpacity) }; break;
                case "K": if (Numbers(operands, 4, out var cmykStroke)) state = state with { Stroke = WithOpacity(Cmyk(cmykStroke[0], cmykStroke[1], cmykStroke[2], cmykStroke[3]), state.StrokeOpacity) }; break;
                case "k": if (Numbers(operands, 4, out var cmykFill)) state = state with { Fill = WithOpacity(Cmyk(cmykFill[0], cmykFill[1], cmykFill[2], cmykFill[3]), state.FillOpacity) }; break;
                case "gs":
                    if (operands.LastOrDefault().Kind == PdfContentKind.Name && resources.TryGetGraphicsState(operands[^1].Text, out var extState) && extState is not null)
                        state = state with
                        {
                            StrokeOpacity = extState.StrokeOpacity,
                            FillOpacity = extState.FillOpacity,
                            Stroke = WithOpacity(state.Stroke, extState.StrokeOpacity),
                            Fill = WithOpacity(state.Fill, extState.FillOpacity)
                        };
                    break;
                case "m":
                    if (Numbers(operands, 2, out var move)) currentPoint = subpathStart = state.Transform.Transform(move[0], move[1]);
                    break;
                case "l":
                    if (currentPoint is { } start && Numbers(operands, 2, out var line))
                    {
                        var end = state.Transform.Transform(line[0], line[1]);
                        path.Add(new PdfLineSegment(start.X, start.Y, end.X, end.Y, state.LineWidth, state.Stroke));
                        currentPoint = end;
                    }
                    break;
                case "re":
                    if (Numbers(operands, 4, out var rectangle))
                    {
                        var a = state.Transform.Transform(rectangle[0], rectangle[1]);
                        var b = state.Transform.Transform(rectangle[0] + rectangle[2], rectangle[1]);
                        var c = state.Transform.Transform(rectangle[0] + rectangle[2], rectangle[1] + rectangle[3]);
                        var d = state.Transform.Transform(rectangle[0], rectangle[1] + rectangle[3]);
                        path.Add(new PdfLineSegment(a.X, a.Y, b.X, b.Y, state.LineWidth, state.Stroke));
                        path.Add(new PdfLineSegment(b.X, b.Y, c.X, c.Y, state.LineWidth, state.Stroke));
                        path.Add(new PdfLineSegment(c.X, c.Y, d.X, d.Y, state.LineWidth, state.Stroke));
                        path.Add(new PdfLineSegment(d.X, d.Y, a.X, a.Y, state.LineWidth, state.Stroke));
                        fillPath.Add(new PdfFilledPolygon([a, b, c, d], state.Fill));
                        currentPoint = null;
                    }
                    break;
                case "h":
                    if (currentPoint is { } point && subpathStart is { } first)
                        path.Add(new PdfLineSegment(point.X, point.Y, first.X, first.Y, state.LineWidth, state.Stroke));
                    currentPoint = subpathStart;
                    break;
                case "S": case "s":
                    foreach (var pathLine in path) output.AddLine(pathLine);
                    path.Clear(); fillPath.Clear(); currentPoint = subpathStart = null;
                    break;
                case "B": case "B*": case "b": case "b*":
                    foreach (var pathFill in fillPath) output.AddFill(pathFill);
                    foreach (var pathLine in path) output.AddLine(pathLine);
                    path.Clear(); fillPath.Clear(); currentPoint = subpathStart = null;
                    break;
                case "f": case "F": case "f*":
                    foreach (var pathFill in fillPath) output.AddFill(pathFill);
                    path.Clear(); fillPath.Clear(); currentPoint = subpathStart = null;
                    break;
                case "n":
                    path.Clear(); fillPath.Clear(); currentPoint = subpathStart = null;
                    break;
                case "BT": textMatrix = textLineMatrix = PdfMatrix.Identity; break;
                case "ET": break;
                case "Tf":
                    if (operands.Count >= 2 && operands[^2].Kind == PdfContentKind.Name && operands[^1].TryNumber(out var size))
                    {
                        fontKey = operands[^2].Text;
                        fontSize = Math.Max(1, size);
                    }
                    break;
                case "Tc": if (Number(operands, 1, out var spacing)) characterSpacing = spacing; break;
                case "Tw": if (Number(operands, 1, out var word)) wordSpacing = word; break;
                case "Tz": if (Number(operands, 1, out var scale)) horizontalScaling = scale / 100d; break;
                case "TL": if (Number(operands, 1, out var textLeading)) leading = textLeading; break;
                case "Td":
                    if (Numbers(operands, 2, out var td)) textMatrix = textLineMatrix = textLineMatrix.Translate(td[0], td[1]);
                    break;
                case "TD":
                    if (Numbers(operands, 2, out var tdLeading)) { leading = -tdLeading[1]; textMatrix = textLineMatrix = textLineMatrix.Translate(tdLeading[0], tdLeading[1]); }
                    break;
                case "Tm": if (Numbers(operands, 6, out var tm)) textMatrix = textLineMatrix = new PdfMatrix(tm[0], tm[1], tm[2], tm[3], tm[4], tm[5]); break;
                case "T*": textMatrix = textLineMatrix = textLineMatrix.Translate(0, -leading); break;
                case "Tj": ShowText(operands.LastOrDefault(), ref textMatrix, fontKey, fontSize, characterSpacing, wordSpacing, horizontalScaling, state, output); break;
                case "TJ":
                    if (operands.LastOrDefault().Kind == PdfContentKind.Array)
                        foreach (var value in operands[^1].Array!)
                            if (value.Kind is PdfContentKind.Hex or PdfContentKind.String) ShowText(value, ref textMatrix, fontKey, fontSize, characterSpacing, wordSpacing, horizontalScaling, state, output);
                            else if (value.TryNumber(out var adjustment)) textMatrix = textMatrix.Translate(-adjustment / 1000d * fontSize * horizontalScaling, 0);
                    break;
                case "'": textMatrix = textLineMatrix = textLineMatrix.Translate(0, -leading); ShowText(operands.LastOrDefault(), ref textMatrix, fontKey, fontSize, characterSpacing, wordSpacing, horizontalScaling, state, output); break;
                case "\"":
                    if (operands.Count >= 3 && operands[^3].TryNumber(out var quoteWordSpacing) && operands[^2].TryNumber(out var quoteCharacterSpacing))
                    {
                        wordSpacing = quoteWordSpacing;
                        characterSpacing = quoteCharacterSpacing;
                        textMatrix = textLineMatrix = textLineMatrix.Translate(0, -leading);
                        ShowText(operands[^1], ref textMatrix, fontKey, fontSize, characterSpacing, wordSpacing, horizontalScaling, state, output);
                    }
                    break;
                case "Do":
                    if (operands.LastOrDefault().Kind == PdfContentKind.Name)
                    {
                        var name = operands[^1].Text;
                        if (resources.TryGetImage(name, out var image) && image is not null)
                            output.AddImage(new PdfPlacedImage(parser.GetImagePayload(source, image), state.Transform, state.FillOpacity, isWatermark));
                        else if (resources.TryGetForm(name, out var form) && form is not null)
                            new PdfContentInterpreter(parser, source, form.Resources).InterpretInto(form.Content, state with { Transform = state.Transform.Multiply(form.Matrix) }, output, depth + 1, isWatermark);
                    }
                    break;
            }
            operands.Clear();
        }
    }

    private void ShowText(PdfContentValue value, ref PdfMatrix textMatrix, string fontKey, double fontSize, double characterSpacing, double wordSpacing, double horizontalScaling, PdfGraphicsState state, PdfPageCommands output)
    {
        if (value.Kind is not (PdfContentKind.Hex or PdfContentKind.String) || value.Bytes is null) return;
        resources.TryGetFont(fontKey, out var font);
        var decoded = font?.DecodeWithGlyphs(value.Bytes) ?? PdfFont.DecodeIdentityWithGlyphs(value.Bytes);
        if (string.IsNullOrEmpty(decoded.Text)) return;
        var transform = state.Transform.Multiply(textMatrix);
        var position = transform.Transform(0, 0);
        // Office/WPS exports commonly use a 0.05 text matrix together with a
        // 209pt font. Do not clamp that matrix to 0.1, otherwise every glyph
        // becomes twice as large and overlaps the next explicitly positioned
        // glyph. The tiny floor only protects a malformed zero-scale matrix.
        var scale = Math.Max(.001, transform.VerticalScale);
        var displaySize = fontSize * scale;
        output.AddText(new PdfTextSegment(decoded.Text, position.X, position.Y, displaySize, font?.Family, state.Fill));
        var advance = decoded.Glyphs.Sum(glyph => (font?.GlyphWidth(glyph) ?? PdfFont.EstimateWidth(glyph.Text)) / 1000d * fontSize + characterSpacing + (glyph.Text == " " ? wordSpacing : 0));
        textMatrix = textMatrix.Translate(advance * horizontalScaling, 0);
    }

    private static bool Number(IReadOnlyList<PdfContentValue> values, int count, out double number)
    {
        number = 0;
        return values.Count >= count && values[^count].TryNumber(out number);
    }

    private static bool Numbers(IReadOnlyList<PdfContentValue> values, int count, out double[] result)
    {
        result = [];
        if (values.Count < count) return false;
        result = new double[count];
        for (var index = 0; index < count; index++)
            if (!values[values.Count - count + index].TryNumber(out result[index])) return false;
        return true;
    }

    private static Color Rgb(double red, double green, double blue) => Color.FromRgb(ToByte(red), ToByte(green), ToByte(blue));

    private static Color Gray(double value) => Rgb(value, value, value);

    private static Color Cmyk(double cyan, double magenta, double yellow, double black) => Rgb(
        1 - Math.Min(1, cyan + black),
        1 - Math.Min(1, magenta + black),
        1 - Math.Min(1, yellow + black));

    private static Color WithOpacity(Color color, double opacity) => Color.FromArgb(
        (byte)Math.Clamp((int)Math.Round(opacity * 255), 0, 255),
        color.R,
        color.G,
        color.B);

    private static byte ToByte(double value) => (byte)Math.Clamp((int)Math.Round(value * 255), 0, 255);
}

internal sealed class PdfPageCommands
{
    public List<PdfTextSegment> Text { get; } = [];
    public List<PdfLineSegment> Lines { get; } = [];
    public List<PdfFilledPolygon> Fills { get; } = [];
    public List<PdfPlacedImage> Images { get; } = [];
    public List<PdfDrawOperation> DrawOperations { get; } = [];

    public void AddText(PdfTextSegment text) { Text.Add(text); DrawOperations.Add(new PdfTextDrawOperation(text)); }
    public void AddLine(PdfLineSegment line) { Lines.Add(line); DrawOperations.Add(new PdfLineDrawOperation(line)); }
    public void AddFill(PdfFilledPolygon fill) { Fills.Add(fill); DrawOperations.Add(new PdfFillDrawOperation(fill)); }
    public void AddImage(PdfPlacedImage image) { Images.Add(image); DrawOperations.Add(new PdfImageDrawOperation(image)); }
}

internal abstract record PdfDrawOperation;
internal sealed record PdfTextDrawOperation(PdfTextSegment Text) : PdfDrawOperation;
internal sealed record PdfLineDrawOperation(PdfLineSegment Line) : PdfDrawOperation;
internal sealed record PdfFillDrawOperation(PdfFilledPolygon Fill) : PdfDrawOperation;
internal sealed record PdfImageDrawOperation(PdfPlacedImage Image) : PdfDrawOperation;
internal sealed record PdfPlacedImage(PdfImagePayload Payload, PdfMatrix Transform, double Opacity, bool IsWatermark);
internal sealed record PdfDecodedGlyph(int Code, string Text);

internal sealed record PdfDecodedText(string Text, IReadOnlyList<PdfDecodedGlyph> Glyphs);

internal sealed record PdfFont(
    string? Family,
    IReadOnlyDictionary<string, string> UnicodeMap,
    IReadOnlyDictionary<int, double> Widths,
    double DefaultWidth,
    int DefaultCodeBytes)
{
    public PdfDecodedText DecodeWithGlyphs(byte[] bytes)
    {
        var keyWidths = UnicodeMap.Keys
            .Where(item => item.Length > 0 && item.Length % 2 == 0)
            .Select(item => item.Length / 2)
            .Distinct()
            .OrderByDescending(item => item)
            .ToArray();
        if (keyWidths.Length == 0) keyWidths = [DefaultCodeBytes];

        var glyphs = new List<PdfDecodedGlyph>();
        for (var offset = 0; offset < bytes.Length;)
        {
            string? mapped = null;
            var used = 0;
            foreach (var width in keyWidths)
            {
                if (offset + width > bytes.Length) continue;
                var key = Convert.ToHexString(bytes, offset, width);
                if (!UnicodeMap.TryGetValue(key, out mapped)) continue;
                used = width;
                break;
            }
            if (mapped is null)
            {
                used = Math.Min(keyWidths[^1], bytes.Length - offset);
                mapped = DecodeIdentity(bytes.AsSpan(offset, used).ToArray());
            }
            var code = 0;
            for (var index = 0; index < used; index++) code = (code << 8) | bytes[offset + index];
            glyphs.Add(new PdfDecodedGlyph(code, mapped));
            offset += used;
        }
        return new PdfDecodedText(string.Concat(glyphs.Select(glyph => glyph.Text)), glyphs);
    }

    public static string DecodeIdentity(byte[] bytes) => bytes.Length % 2 == 0 && bytes.Length > 1
        ? Encoding.BigEndianUnicode.GetString(bytes)
        : Encoding.Latin1.GetString(bytes);

    public static PdfDecodedText DecodeIdentityWithGlyphs(byte[] bytes)
    {
        var codeBytes = bytes.Length > 1 && bytes.Length % 2 == 0 ? 2 : 1;
        var glyphs = new List<PdfDecodedGlyph>();
        for (var offset = 0; offset < bytes.Length; offset += codeBytes)
        {
            var length = Math.Min(codeBytes, bytes.Length - offset);
            var code = 0;
            for (var index = 0; index < length; index++) code = (code << 8) | bytes[offset + index];
            glyphs.Add(new PdfDecodedGlyph(code, DecodeIdentity(bytes.AsSpan(offset, length).ToArray())));
        }
        return new PdfDecodedText(string.Concat(glyphs.Select(glyph => glyph.Text)), glyphs);
    }

    public double GlyphWidth(PdfDecodedGlyph glyph) => Widths.TryGetValue(glyph.Code, out var width) ? width : DefaultWidth;

    public static double EstimateWidth(string text) => text.All(character => character <= 0x7F) ? 500 : 1000;

    public static double EstimateAdvance(string text, double fontSize)
    {
        var advance = 0d;
        foreach (var character in text) advance += character <= 0x7F ? .5 : 1;
        return advance * fontSize;
    }
}

internal readonly record struct PdfPoint(double X, double Y);
internal readonly record struct PdfMatrix(double A, double B, double C, double D, double E, double F)
{
    public static PdfMatrix Identity { get; } = new(1, 0, 0, 1, 0, 0);
    public double VerticalScale => Math.Sqrt(C * C + D * D);
    public PdfPoint Transform(double x, double y) => new(A * x + C * y + E, B * x + D * y + F);
    public PdfMatrix Multiply(PdfMatrix right) => new(
        A * right.A + C * right.B,
        B * right.A + D * right.B,
        A * right.C + C * right.D,
        B * right.C + D * right.D,
        A * right.E + C * right.F + E,
        B * right.E + D * right.F + F);
    public PdfMatrix Translate(double x, double y) => Multiply(new PdfMatrix(1, 0, 0, 1, x, y));
}

internal readonly record struct PdfGraphicsState(PdfMatrix Transform, Color Stroke, Color Fill, double LineWidth, double StrokeOpacity, double FillOpacity)
{
    public static PdfGraphicsState Default { get; } = new(PdfMatrix.Identity, Colors.Black, Colors.Black, 1, 1, 1);
}

internal sealed class PdfSource
{
    private readonly Dictionary<int, PdfRawObject> _objects;
    public IReadOnlyDictionary<int, PdfRawObject> Objects => _objects;

    private PdfSource(Dictionary<int, PdfRawObject> objects) => _objects = objects;

    public static PdfSource Read(string path)
    {
        return Read(File.ReadAllBytes(path));
    }

    public static PdfSource Read(byte[] bytes)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        var objects = new Dictionary<int, PdfRawObject>();
        var position = 0;
        while (position < bytes.Length)
        {
            if (!TryObjectHeader(bytes, position, out var number, out var bodyStart)) { position++; continue; }
            var end = FindObjectEnd(bytes, bodyStart);
            if (end < 0) break;
            var body = bytes[bodyStart..end];
            var reader = new PdfSyntax(body);
            var value = reader.ReadValue();
            var dictionary = value as PdfDictionary;
            objects[number] = new PdfRawObject(number, body, value, dictionary, reader.Position);
            position = end + 6;
        }
        if (objects.Count == 0) throw new InvalidDataException("The file does not contain readable PDF objects.");
        var source = new PdfSource(objects);
        source.ExpandObjectStreams();
        return source;
    }

    private static int FindObjectEnd(byte[] data, int bodyStart)
    {
        // A stream's binary payload may coincidentally contain "endobj". When
        // the stream length is direct (the usual case for images and page
        // content), jump over it instead of scanning every byte. This keeps an
        // 88 MB, image-heavy PDF linear and avoids false object boundaries.
        var probeLength = Math.Min(4_096, data.Length - bodyStart);
        if (probeLength > 0)
        {
            var probe = data[bodyStart..(bodyStart + probeLength)];
            var reader = new PdfSyntax(probe);
            if (reader.ReadValue() is PdfDictionary dictionary && dictionary.Get("Length") is PdfNumber lengthValue)
            {
                var stream = IndexOf(data, "stream", bodyStart + reader.Position);
                if (stream >= 0 && stream - (bodyStart + reader.Position) <= 64)
                {
                    var streamStart = stream + 6;
                    if (streamStart < data.Length && data[streamStart] == '\r') streamStart++;
                    if (streamStart < data.Length && data[streamStart] == '\n') streamStart++;
                    var streamEnd = streamStart + (long)lengthValue.Value;
                    if (streamEnd is >= 0 and < int.MaxValue && streamEnd <= data.Length)
                    {
                        var afterStream = (int)streamEnd;
                        while (afterStream < data.Length && PdfSyntax.IsWhitespace(data[afterStream])) afterStream++;
                        if (Matches(data, afterStream, "endstream"))
                        {
                            afterStream += 9;
                            while (afterStream < data.Length && PdfSyntax.IsWhitespace(data[afterStream])) afterStream++;
                            if (Matches(data, afterStream, "endobj")) return afterStream;
                        }
                    }
                }
            }
        }
        return IndexOf(data, "endobj", bodyStart);
    }

    /// <summary>
    /// PDF 1.5+ is allowed to put ordinary (non-stream) objects into a Flate
    /// compressed /ObjStm. Page trees, resource dictionaries and the catalog
    /// commonly live there, so scanning only physical "n n obj" records makes
    /// an otherwise valid PDF look as if it has no pages.
    /// </summary>
    private void ExpandObjectStreams()
    {
        foreach (var container in _objects.Values
                     .Where(item => string.Equals(NameOf(item.Dictionary?.Get("Type")), "ObjStm", StringComparison.Ordinal))
                     .ToArray())
        {
            var dictionary = container.Dictionary!;
            var count = (int)NumberOf(dictionary.Get("N"));
            var first = (int)NumberOf(dictionary.Get("First"));
            if (count <= 0 || first < 0) continue;

            byte[] data;
            try { data = DecodedStreamOf(container); }
            catch (InvalidDataException) { continue; }
            if (first > data.Length) continue;

            var header = new PdfSyntax(data);
            var entries = new List<(int Number, int Offset)>(count);
            for (var index = 0; index < count; index++)
            {
                if (header.ReadValue() is not PdfNumber objectNumber || header.ReadValue() is not PdfNumber objectOffset) break;
                entries.Add(((int)objectNumber.Value, (int)objectOffset.Value));
            }

            var orderedEntries = entries.OrderBy(entry => entry.Offset).ToArray();
            for (var index = 0; index < orderedEntries.Length; index++)
            {
                var entry = orderedEntries[index];
                if (_objects.ContainsKey(entry.Number)) continue;
                var start = first + entry.Offset;
                var end = index + 1 < orderedEntries.Length ? first + orderedEntries[index + 1].Offset : data.Length;
                if (start < first || start >= end || end > data.Length) continue;
                // Every object previously retained data[start..end-of-object-stream].
                // A 3,000-object stream therefore multiplied memory by thousands.
                var body = data[start..end];
                var reader = new PdfSyntax(body);
                var value = reader.ReadValue();
                if (value is null) continue;
                _objects[entry.Number] = new PdfRawObject(entry.Number, body, value, value as PdfDictionary, reader.Position);
            }
        }
    }

    public PdfRawObject? ObjectOf(PdfValue? value) => value switch
    {
        PdfReference reference when Objects.TryGetValue(reference.Number, out var item) => item,
        _ => null
    };

    public PdfRawObject? ObjectOf(PdfReference? reference) => reference is not null && Objects.TryGetValue(reference.Number, out var item) ? item : null;

    public PdfDictionary? DictionaryOf(PdfValue? value) => value switch
    {
        PdfDictionary dictionary => dictionary,
        PdfReference reference => DictionaryOf(reference),
        _ => null
    };

    public PdfDictionary DictionaryOf(PdfReference reference) => ObjectOf(reference)?.Dictionary ?? throw new InvalidDataException($"PDF object {reference.Number} is missing or not a dictionary.");

    public PdfArray? ArrayOf(PdfValue? value) => value switch
    {
        PdfArray array => array,
        PdfReference reference => ObjectOf(reference)?.Value as PdfArray,
        _ => null
    };

    public string? NameOf(PdfValue? value) => value switch
    {
        PdfName name => name.Value,
        PdfReference reference => ObjectOf(reference)?.Value is PdfName name ? name.Value : null,
        _ => null
    };

    public double NumberOf(PdfValue? value, double fallback = 0) => value switch
    {
        PdfNumber number => (double)number.Value,
        PdfReference reference when ObjectOf(reference)?.Value is PdfNumber number => (double)number.Value,
        _ => fallback
    };

    public IReadOnlyList<string> FilterNames(PdfRawObject item)
    {
        var value = item.Dictionary?.Get("Filter");
        if (value is PdfName name) return [name.Value];
        if (value is PdfArray array) return array.Values.OfType<PdfName>().Select(name => name.Value).ToArray();
        return [];
    }

    public byte[] EncodedStreamOf(PdfRawObject item)
    {
        if (item.Dictionary is null) return [];
        var streamMarker = IndexOf(item.Body, "stream", item.ValueEnd);
        if (streamMarker < 0) return [];
        var start = streamMarker + 6;
        if (start < item.Body.Length && item.Body[start] == '\r') start++;
        if (start < item.Body.Length && item.Body[start] == '\n') start++;
        var length = NumberOf(item.Dictionary.Get("Length"), -1);
        if (length >= 0 && start + length <= item.Body.Length) return item.Body[start..(start + (int)length)];
        var end = IndexOf(item.Body, "endstream", start);
        return end < 0 ? [] : item.Body[start..end].TrimPdfWhitespace();
    }

    public byte[] DecodedStreamOf(PdfRawObject item)
    {
        var data = EncodedStreamOf(item);
        foreach (var filter in FilterNames(item))
        {
            if (filter is "FlateDecode" or "Fl") data = Inflate(data);
            else if (filter is "ASCIIHexDecode" or "AHx") data = PdfSyntax.ToBytes(Encoding.Latin1.GetString(data));
            else throw new NotSupportedException($"PDF stream filter '{filter}' is not supported by the common PDF subset.");
        }
        return data;
    }

    private static byte[] Inflate(byte[] input)
    {
        try { return Inflate(input, stream => new ZLibStream(stream, CompressionMode.Decompress)); }
        catch (InvalidDataException) { return Inflate(input, stream => new DeflateStream(stream, CompressionMode.Decompress)); }
    }

    private static byte[] Inflate(byte[] input, Func<Stream, Stream> create)
    {
        using var source = new MemoryStream(input, writable: false);
        using var compressed = create(source);
        using var output = new MemoryStream();
        compressed.CopyTo(output);
        return output.ToArray();
    }

    private static bool TryObjectHeader(byte[] data, int position, out int number, out int bodyStart)
    {
        number = 0; bodyStart = 0;
        if (position > 0 && data[position - 1] is not ((byte)'\n' or (byte)'\r')) return false;
        var index = position;
        if (!ReadInteger(data, ref index, out number) || !SkipSpace(data, ref index) || !ReadInteger(data, ref index, out _) || !SkipSpace(data, ref index)) return false;
        if (!Matches(data, index, "obj")) return false;
        bodyStart = index + 3;
        return true;
    }

    private static bool ReadInteger(byte[] data, ref int index, out int value)
    {
        value = 0;
        var start = index;
        while (index < data.Length && data[index] is >= (byte)'0' and <= (byte)'9') { value = value * 10 + data[index] - '0'; index++; }
        return index > start;
    }

    private static bool SkipSpace(byte[] data, ref int index)
    {
        var start = index;
        while (index < data.Length && PdfSyntax.IsWhitespace(data[index])) index++;
        return index > start;
    }

    private static int IndexOf(byte[] data, string needle, int start)
    {
        var bytes = Encoding.ASCII.GetBytes(needle);
        for (var index = start; index <= data.Length - bytes.Length; index++)
        {
            var matches = true;
            for (var part = 0; part < bytes.Length; part++) if (data[index + part] != bytes[part]) { matches = false; break; }
            if (matches) return index;
        }
        return -1;
    }

    private static bool Matches(byte[] data, int start, string text)
    {
        if (start + text.Length > data.Length) return false;
        for (var index = 0; index < text.Length; index++) if (data[start + index] != (byte)text[index]) return false;
        return true;
    }
}

internal sealed class PdfRawObject(int number, byte[] body, PdfValue? value, PdfDictionary? dictionary, int valueEnd)
{
    public int Number { get; } = number;
    public byte[] Body { get; } = body;
    public PdfDictionary? Dictionary { get; } = dictionary;
    public int ValueEnd { get; } = valueEnd;
    public PdfValue? Value { get; } = value;
}

internal abstract class PdfValue;
internal sealed class PdfName(string value) : PdfValue { public string Value { get; } = value; }
internal sealed class PdfNumber(decimal value) : PdfValue { public decimal Value { get; } = value; }
internal sealed class PdfReference(int number) : PdfValue { public int Number { get; } = number; }
internal sealed class PdfString(byte[] bytes) : PdfValue { public byte[] Bytes { get; } = bytes; }
internal sealed class PdfArray(List<PdfValue> values) : PdfValue { public IReadOnlyList<PdfValue> Values { get; } = values; }
internal sealed class PdfDictionary(Dictionary<string, PdfValue> values) : PdfValue
{
    public IReadOnlyDictionary<string, PdfValue> Values { get; } = values;
    public PdfValue? Get(string key) => values.TryGetValue(key, out var value) ? value : null;
}
internal sealed class PdfKeyword(string value) : PdfValue { public string Value { get; } = value; }

internal sealed class PdfSyntax
{
    private readonly byte[] _data;
    public int Position { get; private set; }
    public PdfSyntax(byte[] data) => _data = data;

    public PdfValue? ReadValue()
    {
        SkipWhitespaceAndComments();
        if (Position >= _data.Length) return null;
        if (StartsWith("<<")) return ReadDictionary();
        return _data[Position] switch
        {
            (byte)'/' => new PdfName(ReadName()),
            (byte)'[' => ReadArray(),
            (byte)'<' => new PdfString(ReadHex()),
            (byte)'(' => new PdfString(ReadLiteral()),
            _ => ReadNumberReferenceOrKeyword()
        };
    }

    private PdfDictionary ReadDictionary()
    {
        Position += 2;
        var values = new Dictionary<string, PdfValue>(StringComparer.Ordinal);
        while (Position < _data.Length)
        {
            SkipWhitespaceAndComments();
            if (StartsWith(">>")) { Position += 2; break; }
            if (Position >= _data.Length || _data[Position] != '/') break;
            var key = ReadName();
            var value = ReadValue();
            if (value is not null) values[key] = value;
        }
        return new PdfDictionary(values);
    }

    private PdfArray ReadArray()
    {
        Position++;
        var values = new List<PdfValue>();
        while (Position < _data.Length)
        {
            SkipWhitespaceAndComments();
            if (Position < _data.Length && _data[Position] == ']') { Position++; break; }
            var value = ReadValue();
            if (value is null) break;
            values.Add(value);
        }
        return new PdfArray(values);
    }

    private PdfValue ReadNumberReferenceOrKeyword()
    {
        var first = ReadBareWord();
        if (!decimal.TryParse(first, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)) return new PdfKeyword(first);
        var checkpoint = Position;
        SkipWhitespaceAndComments();
        var second = ReadBareWord();
        if (decimal.TryParse(second, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
        {
            SkipWhitespaceAndComments();
            var marker = ReadBareWord();
            if (marker == "R") return new PdfReference((int)number);
        }
        Position = checkpoint;
        return new PdfNumber(number);
    }

    private string ReadName()
    {
        Position++;
        var buffer = new List<byte>();
        while (Position < _data.Length && !IsDelimiter(_data[Position]))
        {
            if (_data[Position] == '#' && Position + 2 < _data.Length)
            {
                var pair = Encoding.ASCII.GetString(_data, Position + 1, 2);
                if (byte.TryParse(pair, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var decoded)) { buffer.Add(decoded); Position += 3; continue; }
            }
            buffer.Add(_data[Position++]);
        }
        return Encoding.Latin1.GetString([.. buffer]);
    }

    private byte[] ReadHex()
    {
        Position++;
        var text = new StringBuilder();
        while (Position < _data.Length && _data[Position] != '>')
        {
            if (!IsWhitespace(_data[Position])) text.Append((char)_data[Position]);
            Position++;
        }
        if (Position < _data.Length) Position++;
        return ToBytes(text.ToString());
    }

    private byte[] ReadLiteral()
    {
        Position++;
        var output = new List<byte>();
        var depth = 1;
        while (Position < _data.Length && depth > 0)
        {
            var value = _data[Position++];
            if (value == '\\' && Position < _data.Length)
            {
                var escaped = _data[Position++];
                output.Add(escaped switch { (byte)'n' => (byte)'\n', (byte)'r' => (byte)'\r', (byte)'t' => (byte)'\t', (byte)'b' => (byte)'\b', (byte)'f' => (byte)'\f', _ => escaped });
            }
            else if (value == '(') { depth++; output.Add(value); }
            else if (value == ')') { depth--; if (depth > 0) output.Add(value); }
            else output.Add(value);
        }
        return [.. output];
    }

    private string ReadBareWord()
    {
        SkipWhitespaceAndComments();
        var start = Position;
        while (Position < _data.Length && !IsDelimiter(_data[Position])) Position++;
        return Encoding.Latin1.GetString(_data, start, Position - start);
    }

    private void SkipWhitespaceAndComments()
    {
        while (Position < _data.Length)
        {
            if (IsWhitespace(_data[Position])) { Position++; continue; }
            if (_data[Position] != '%') break;
            while (Position < _data.Length && _data[Position] is not ((byte)'\n' or (byte)'\r')) Position++;
        }
    }

    private bool StartsWith(string text)
    {
        if (Position + text.Length > _data.Length) return false;
        for (var index = 0; index < text.Length; index++) if (_data[Position + index] != text[index]) return false;
        return true;
    }

    public static bool IsWhitespace(byte value) => value is 0 or 9 or 10 or 12 or 13 or 32;
    private static bool IsDelimiter(byte value) => IsWhitespace(value) || value is (byte)'/' or (byte)'[' or (byte)']' or (byte)'<' or (byte)'>' or (byte)'(' or (byte)')' or (byte)'%';
    public static byte[] ToBytes(string text)
    {
        if (text.Length % 2 != 0) text += "0";
        var output = new byte[text.Length / 2];
        for (var index = 0; index < output.Length; index++) output[index] = byte.TryParse(text.AsSpan(index * 2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var value) ? value : (byte)0;
        return output;
    }
}

internal enum PdfContentKind { Number, Name, Hex, String, Array, Dictionary, Word }
internal readonly record struct PdfContentValue(PdfContentKind Kind, string Text, byte[]? Bytes, IReadOnlyList<PdfContentValue>? Array, double Numeric)
{
    public bool TryNumber(out double value) { value = Numeric; return Kind == PdfContentKind.Number; }
}

internal sealed class PdfContentReader
{
    private readonly byte[] _data;
    private int _position;
    public PdfContentReader(byte[] data) => _data = data;

    public bool TryRead(out PdfContentValue value)
    {
        Skip();
        if (_position >= _data.Length) { value = default; return false; }
        var current = _data[_position];
        if (current == '/') { value = new(PdfContentKind.Name, ReadName(), null, null, 0); return true; }
        if (current == '[') { value = new(PdfContentKind.Array, string.Empty, null, ReadArray(), 0); return true; }
        if (current == '<' && NextIs('<')) { value = new(PdfContentKind.Dictionary, ReadDictionaryText(), null, null, 0); return true; }
        if (current == '<' && !NextIs('<')) { value = new(PdfContentKind.Hex, string.Empty, ReadHex(), null, 0); return true; }
        if (current == '(') { value = new(PdfContentKind.String, string.Empty, ReadLiteral(), null, 0); return true; }
        var word = ReadWord();
        value = double.TryParse(word, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
            ? new PdfContentValue(PdfContentKind.Number, word, null, null, number)
            : new PdfContentValue(PdfContentKind.Word, word, null, null, 0);
        return word.Length > 0;
    }

    private IReadOnlyList<PdfContentValue> ReadArray()
    {
        _position++;
        var result = new List<PdfContentValue>();
        while (true)
        {
            Skip();
            if (_position >= _data.Length || _data[_position] == ']') { if (_position < _data.Length) _position++; break; }
            if (!TryRead(out var value)) break;
            result.Add(value);
        }
        return result;
    }

    private string ReadName()
    {
        _position++;
        var start = _position;
        while (_position < _data.Length && !Delimiter(_data[_position])) _position++;
        return Encoding.Latin1.GetString(_data, start, _position - start);
    }

    private byte[] ReadHex()
    {
        _position++;
        var builder = new StringBuilder();
        while (_position < _data.Length && _data[_position] != '>') { if (!PdfSyntax.IsWhitespace(_data[_position])) builder.Append((char)_data[_position]); _position++; }
        if (_position < _data.Length) _position++;
        return PdfSyntax.ToBytes(builder.ToString());
    }

    private byte[] ReadLiteral()
    {
        _position++;
        var result = new List<byte>();
        var depth = 1;
        while (_position < _data.Length && depth > 0)
        {
            var current = _data[_position++];
            if (current == '\\' && _position < _data.Length) result.Add(_data[_position++]);
            else if (current == '(') { depth++; result.Add(current); }
            else if (current == ')') { depth--; if (depth > 0) result.Add(current); }
            else result.Add(current);
        }
        return [.. result];
    }

    private string ReadDictionaryText()
    {
        var start = _position;
        _position += 2;
        var depth = 1;
        while (_position < _data.Length && depth > 0)
        {
            if (_data[_position] == '(') { ReadLiteral(); continue; }
            if (_data[_position] == '<' && NextIs('<')) { _position += 2; depth++; continue; }
            if (_data[_position] == '>' && NextIs('>')) { _position += 2; depth--; continue; }
            _position++;
        }
        return Encoding.Latin1.GetString(_data, start, _position - start);
    }

    private string ReadWord()
    {
        var start = _position;
        while (_position < _data.Length && !Delimiter(_data[_position])) _position++;
        return Encoding.Latin1.GetString(_data, start, _position - start);
    }

    private bool NextIs(char value) => _position + 1 < _data.Length && _data[_position + 1] == value;
    private void Skip() { while (_position < _data.Length && PdfSyntax.IsWhitespace(_data[_position])) _position++; }
    private static bool Delimiter(byte value) => PdfSyntax.IsWhitespace(value) || value is (byte)'/' or (byte)'[' or (byte)']' or (byte)'<' or (byte)'>' or (byte)'(' or (byte)')';
}

internal static class PdfByteExtensions
{
    public static byte[] TrimPdfWhitespace(this byte[] value)
    {
        var start = 0;
        var end = value.Length;
        while (start < end && PdfSyntax.IsWhitespace(value[start])) start++;
        while (end > start && PdfSyntax.IsWhitespace(value[end - 1])) end--;
        return start == 0 && end == value.Length ? value : value[start..end];
    }
}
