using System;
using System.Collections.Generic;
using Avalonia;
using global::Avalonia.Layout;
using global::Avalonia.Media;

namespace OfficeViewer.Avalonia.Pptx;

internal sealed class PptxDocumentModel
{
    // A presentation loaded through the public Document byte[] property keeps the
    // package bytes here so pictures can still be decoded lazily. The viewer drops
    // this reference together with the visual tree when Document becomes null.
    public string? SourcePath { get; init; }
    public byte[]? SourceBytes { get; init; }
    public double Width { get; init; } = 1280;
    public double Height { get; init; } = 720;
    public List<PptxSlideModel> Slides { get; } = [];
}

internal sealed class PptxSlideModel
{
    public PptxFill? Background { get; init; }
    public List<PptxElement> Elements { get; } = [];
}

internal abstract class PptxElement
{
    public double Left { get; init; }
    public double Top { get; init; }
    public double Width { get; init; }
    public double Height { get; init; }
    public double Rotation { get; init; }
    public bool FlipHorizontal { get; init; }
    public bool FlipVertical { get; init; }
}

internal sealed class PptxShape : PptxElement
{
    public string Geometry { get; init; } = "rect";
    public PptxFill? Fill { get; init; }
    public PptxFill? Stroke { get; init; }
    public double StrokeThickness { get; init; }
    public List<PptxParagraph> Paragraphs { get; } = [];
    public Thickness TextInsets { get; init; }
    public TextAlignment TextAlignment { get; init; } = TextAlignment.Left;
    public VerticalAlignment VerticalAlignment { get; init; } = VerticalAlignment.Top;
}

internal sealed class PptxPicture : PptxElement
{
    public required string PackagePath { get; init; }
}

internal sealed class PptxLine : PptxElement
{
    public PptxFill? Stroke { get; init; }
    public double StrokeThickness { get; init; } = 1;
}

internal sealed class PptxParagraph
{
    public List<PptxTextRun> Runs { get; } = [];
    public TextAlignment? Alignment { get; init; }
    public double? SpaceAfter { get; init; }
}

internal sealed class PptxTextRun
{
    public required string Text { get; init; }
    public string? FontFamily { get; init; }
    public double? FontSize { get; init; }
    public FontWeight? FontWeight { get; init; }
    public FontStyle? FontStyle { get; init; }
    public PptxFill? Foreground { get; init; }
    public bool Underline { get; init; }
}

internal sealed class PptxFill
{
    public string? Color { get; init; }
    public List<PptxGradientStop> GradientStops { get; } = [];
    public double GradientAngle { get; init; }
}

internal readonly record struct PptxGradientStop(double Position, string Color);

internal sealed class PptxChart : PptxElement
{
    public string? Title { get; init; }
    public List<string> Categories { get; } = [];
    public List<double> Values { get; } = [];
    public PptxFill? BarFill { get; init; }
}

internal readonly record struct PptxTransform(double OffsetX, double OffsetY, double ScaleX, double ScaleY)
{
    public static PptxTransform Identity => new(0, 0, 1, 1);
    public double X(double value) => OffsetX + value * ScaleX;
    public double Y(double value) => OffsetY + value * ScaleY;
}

internal readonly record struct PptxBounds(double Left, double Top, double Width, double Height, double Rotation, bool FlipHorizontal, bool FlipVertical);
