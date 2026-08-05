# OfficeViewer.Avalonia

[中文](#中文说明) | [English](#english)

## 中文说明

`OfficeViewer.Avalonia` 是用于 Avalonia 的只读 Office/PDF 原生预览控件包，提供 DOCX、PPTX、XLSX 与常见 PDF 的连续流式预览。

控件外层保持透明，不会强制宿主窗口背景。请由应用在外层使用 `Border`、`Panel` 或主题容器决定背景、圆角与阴影。

### 参考项目

DOCX、PPTX、XLSX 的文件读取、关系解析、样式模型和连续布局设计参考了 [vue-office-core](https://github.com/501351981/vue-office) 项目。本包将对应思路移植为 C# 的 Open XML 解析和 Avalonia 控件渲染；它不包含 Vue、JavaScript 或 WebView 运行时，也不是对 Vue 组件的直接包装。

PDF 预览是当前项目独立实现的纯托管常见子集解析器：它不捆绑 PDF.js、PDFium 或 PdfPig。支持范围请见“格式范围与限制”。

### 包含控件

| 文档格式 | 命名空间和控件 | 读取结果属性 |
| --- | --- | --- |
| Word / DOCX | `OfficeViewer.Avalonia.Docx.DocxViewer` | `PageCount` |
| PowerPoint / PPTX | `OfficeViewer.Avalonia.Pptx.PptxViewer` | `PageCount` |
| Excel / XLSX | `OfficeViewer.Avalonia.Xlsx.XlsxViewer` | `SheetCount`、`SelectedSheetIndex` |
| PDF | `OfficeViewer.Avalonia.Pdf.PdfViewer` | `PageCount` |

每个控件均提供可绑定的 `byte[]? Document` 属性。给 `Document` 赋新的文件字节数组会自动开始读取；将其清除或设为 `null` 会自动取消旧读取任务，并释放当前文档的模型、图像与可视资源。PPTX/PDF 仅会实例化可视页及少量相邻页，避免长文档持续占用图片内存。控件仍保留 `LoadAsync` 作为非绑定场景的便捷方法。

### 缩放

**所有四个预览控件均支持按住 `Ctrl` 并滚动鼠标滚轮进行放大或缩小。** 

### 安装

从你的 NuGet 源安装：

```xml
<PackageReference Include="OfficeViewer.Avalonia" Version="0.1.2" />
```

本仓库本地构建后的包目录如下，可在 Visual Studio、Rider 或 `NuGet.Config` 中添加为本地源：

```text
OfficeViewer.Avalonia.NuGet\artifacts\packages
```

### Avalonia XAML 用法

```xml
<Window xmlns="https://github.com/avaloniaui"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:docx="clr-namespace:OfficeViewer.Avalonia.Docx;assembly=OfficeViewer.Avalonia"
        xmlns:pptx="clr-namespace:OfficeViewer.Avalonia.Pptx;assembly=OfficeViewer.Avalonia"
        xmlns:xlsx="clr-namespace:OfficeViewer.Avalonia.Xlsx;assembly=OfficeViewer.Avalonia"
        xmlns:pdf="clr-namespace:OfficeViewer.Avalonia.Pdf;assembly=OfficeViewer.Avalonia">
  <Grid RowDefinitions="Auto,*">
    <TextBlock Text="{Binding PageCount, ElementName=DocxPreview}" />
    <Border Grid.Row="1" Background="White" Padding="8">
      <docx:DocxViewer x:Name="DocxPreview"
                       Document="{Binding DocxDocument}" />
    </Border>
  </Grid>
</Window>
```

其他三种控件的声明方式：

```xml
<pptx:PptxViewer x:Name="PptxPreview"
                 Document="{Binding PptxDocument}" />
<pdf:PdfViewer x:Name="PdfPreview"
               Document="{Binding PdfDocument}" />
<xlsx:XlsxViewer x:Name="XlsxPreview"
                 Document="{Binding XlsxDocument}"
                 SelectedSheetIndex="{Binding SelectedWorksheet, Mode=TwoWay}" />
```

`SelectedSheetIndex` 从 `0` 开始。XLSX 控件故意不内置工作表选择器，请绑定到你自己的按钮、`ComboBox`、选项卡或 ViewModel。工作簿读取完成后，`SheetCount` 会更新。

### C# 加载与释放

```csharp
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

public sealed class OfficePreviewViewModel : INotifyPropertyChanged
{
    private byte[]? _docxDocument;
    private byte[]? _pptxDocument;
    private byte[]? _xlsxDocument;
    private byte[]? _pdfDocument;
    private int _selectedWorksheet;

    public byte[]? DocxDocument { get => _docxDocument; set => SetField(ref _docxDocument, value); }
    public byte[]? PptxDocument { get => _pptxDocument; set => SetField(ref _pptxDocument, value); }
    public byte[]? XlsxDocument { get => _xlsxDocument; set => SetField(ref _xlsxDocument, value); }
    public byte[]? PdfDocument { get => _pdfDocument; set => SetField(ref _pdfDocument, value); }
    public int SelectedWorksheet { get => _selectedWorksheet; set => SetField(ref _selectedWorksheet, value); }

    public async Task OpenDocxAsync(string path) => DocxDocument = await File.ReadAllBytesAsync(path);
    public async Task OpenPptxAsync(string path) => PptxDocument = await File.ReadAllBytesAsync(path);
    public async Task OpenXlsxAsync(string path) => XlsxDocument = await File.ReadAllBytesAsync(path);
    public async Task OpenPdfAsync(string path) => PdfDocument = await File.ReadAllBytesAsync(path);

    public void ReleaseDocuments()
    {
        // null automatically cancels loading and releases each viewer's resources.
        DocxDocument = null;
        PptxDocument = null;
        XlsxDocument = null;
        PdfDocument = null;
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void SetField<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return;
        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
```

`Document`、`PageCount`、`SheetCount` 是 Avalonia StyledProperty。请在 ViewModel 中将文档字节数组绑定到 `Document`；需要释放内存时也将绑定源设为 `null`，使控件和 ViewModel 都不再保留该字节数组。

可在代码中调用 `LoadAsync(Stream)` 读取流。控件不会关闭调用方传入的流，而会创建自己的字节快照并将其赋给 `Document`；MVVM 绑定仍建议使用 `byte[]? Document`。

### NativeAOT

本包支持 NativeAOT。

### 格式范围与限制

- DOCX：WordprocessingML 段落、文字 Run、编号、图片、表格、常见形状/文本框和页面流式布局。
- PPTX：PresentationML 幻灯片、主题/样式继承、文字、图片、填充、渐变、常见形状及已支持图表的连续流布局。
- XLSX：工作簿/工作表关系、共享字符串/富文本、样式、合并单元格、图片、行列尺寸、颜色和当前工作表绑定。
- PDF：压缩对象流、Form XObject、文字/CMap 解码、带常见蒙版的图片、填充/线条/透明度，以及 WPS 水印标记。加密、签名、损坏文件或高级 PDF 图形/字体仍可能需要完整 PDF 渲染器，本包不会宣称完全兼容。

## English

`OfficeViewer.Avalonia` is a set of read-only native Avalonia viewers for DOCX, PPTX, XLSX, and common PDF files. Documents use continuous scrolling rather than a flip-book UI.

The outer viewer surface is transparent. Hosts own the surrounding background, corner radius, and shadow by wrapping the control in a `Border`, `Panel`, or theme container.

### Reference project

The file-reading flow, relationship parsing, style models, and continuous-layout design for DOCX, PPTX, and XLSX are informed by [vue-office-core](https://github.com/501351981/vue-office). This package ports those ideas to C# Open XML parsing and Avalonia-native rendering. It does not bundle Vue, JavaScript, or a WebView, and it is not a wrapper around Vue components.

The PDF viewer is this project's separate managed common-subset implementation. It does not bundle PDF.js, PDFium, or PdfPig. See “Format scope and limitations” for the supported subset.

### Included controls

| Document | Namespace and control | Result count |
| --- | --- | --- |
| Word / DOCX | `OfficeViewer.Avalonia.Docx.DocxViewer` | `PageCount` |
| PowerPoint / PPTX | `OfficeViewer.Avalonia.Pptx.PptxViewer` | `PageCount` |
| Excel / XLSX | `OfficeViewer.Avalonia.Xlsx.XlsxViewer` | `SheetCount`, `SelectedSheetIndex` |
| PDF | `OfficeViewer.Avalonia.Pdf.PdfViewer` | `PageCount` |

### 预览截图 / Screenshots

以下截图来自 samples 目录中的实际控件运行效果。 / The following screenshots show the viewers running with the sample documents in the `samples` directory.

| 格式 / Format | 预览 / Preview |
| --- | --- |
| Word / DOCX | ![DOCX preview](https://github.com/fvffv/OfficeViewer.Avalonia/blob/main/samples/word.png?raw=true) |
| PowerPoint / PPTX | ![PPTX preview](https://github.com/fvffv/OfficeViewer.Avalonia/blob/main/samples/pptx.png?raw=true) |
| Excel / XLSX | ![XLSX preview](https://github.com/fvffv/OfficeViewer.Avalonia/blob/main/samples/xlsx.png?raw=true) |
| PDF | ![PDF preview](https://github.com/fvffv/OfficeViewer.Avalonia/blob/main/samples/pdf.png?raw=true) |

Every viewer exposes a bindable `byte[]? Document` property. Assign a new document byte array to start loading automatically; clear it or assign `null` to cancel stale loading and release the current document model, images, and visual resources. PPTX and PDF only materialize visible pages/slides plus a small adjacent buffer. `LoadAsync` remains available as a convenience method for non-binding scenarios.

### Zoom

**All four viewers support zooming with `Ctrl` + mouse wheel.** 

### Install

Install from your NuGet source:

```xml
<PackageReference Include="OfficeViewer.Avalonia" Version="0.1.2" />
```

For a local build of this repository, add the following folder as a NuGet source:

```text
OfficeViewer.Avalonia.NuGet\artifacts\packages
```

### Avalonia XAML usage

```xml
<Window xmlns="https://github.com/avaloniaui"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:docx="clr-namespace:OfficeViewer.Avalonia.Docx;assembly=OfficeViewer.Avalonia"
        xmlns:pptx="clr-namespace:OfficeViewer.Avalonia.Pptx;assembly=OfficeViewer.Avalonia"
        xmlns:xlsx="clr-namespace:OfficeViewer.Avalonia.Xlsx;assembly=OfficeViewer.Avalonia"
        xmlns:pdf="clr-namespace:OfficeViewer.Avalonia.Pdf;assembly=OfficeViewer.Avalonia">
  <Grid RowDefinitions="Auto,*">
    <TextBlock Text="{Binding PageCount, ElementName=DocxPreview}" />
    <Border Grid.Row="1" Background="White" Padding="8">
      <docx:DocxViewer x:Name="DocxPreview"
                       Document="{Binding DocxDocument}" />
    </Border>
  </Grid>
</Window>
```

Create the other controls in the same way:

```xml
<pptx:PptxViewer x:Name="PptxPreview"
                 Document="{Binding PptxDocument}" />
<pdf:PdfViewer x:Name="PdfPreview"
               Document="{Binding PdfDocument}" />
<xlsx:XlsxViewer x:Name="XlsxPreview"
                 Document="{Binding XlsxDocument}"
                 SelectedSheetIndex="{Binding SelectedWorksheet, Mode=TwoWay}" />
```

`SelectedSheetIndex` is zero-based. The XLSX control intentionally has no built-in worksheet selector; bind it to your own buttons, `ComboBox`, tabs, or view model. `SheetCount` updates after a successful workbook load.

### Load and release in C#

```csharp
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

public sealed class OfficePreviewViewModel : INotifyPropertyChanged
{
    private byte[]? _docxDocument;
    private byte[]? _pptxDocument;
    private byte[]? _xlsxDocument;
    private byte[]? _pdfDocument;
    private int _selectedWorksheet;

    public byte[]? DocxDocument { get => _docxDocument; set => SetField(ref _docxDocument, value); }
    public byte[]? PptxDocument { get => _pptxDocument; set => SetField(ref _pptxDocument, value); }
    public byte[]? XlsxDocument { get => _xlsxDocument; set => SetField(ref _xlsxDocument, value); }
    public byte[]? PdfDocument { get => _pdfDocument; set => SetField(ref _pdfDocument, value); }
    public int SelectedWorksheet { get => _selectedWorksheet; set => SetField(ref _selectedWorksheet, value); }

    public async Task OpenDocxAsync(string path) => DocxDocument = await File.ReadAllBytesAsync(path);
    public async Task OpenPptxAsync(string path) => PptxDocument = await File.ReadAllBytesAsync(path);
    public async Task OpenXlsxAsync(string path) => XlsxDocument = await File.ReadAllBytesAsync(path);
    public async Task OpenPdfAsync(string path) => PdfDocument = await File.ReadAllBytesAsync(path);

    public void ReleaseDocuments()
    {
        // null automatically cancels loading and releases each viewer's resources.
        DocxDocument = null;
        PptxDocument = null;
        XlsxDocument = null;
        PdfDocument = null;
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void SetField<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return;
        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
```

`Document`, `PageCount`, and `SheetCount` are Avalonia styled properties. Bind the document byte array from a view model to `Document`; set that binding source to `null` when releasing memory so neither the viewer nor the view model retains the bytes.

For code-behind or service code, `LoadAsync(Stream)` accepts a stream without taking ownership of it. The viewer creates its own byte snapshot and assigns it to `Document`; `byte[]? Document` remains the recommended MVVM binding API.

### NativeAOT

This package supports NativeAOT.

### Format scope and limitations

- DOCX: WordprocessingML paragraphs, runs, numbering, images, tables, common shapes/text boxes, and page-flow layout.
- PPTX: PresentationML slides, theme/style inheritance, text, pictures, fills, gradients, common shapes, and supported charts in continuous flow.
- XLSX: workbook/sheet relationships, shared strings/rich text, styles, merged cells, images, row/column sizes, colors, and active-sheet binding.
- PDF: compressed object streams, Form XObjects, text/CMap decoding, images with common masks, fills/lines/transparency, and WPS watermark artifacts. Password-protected, signed, malformed, or advanced PDF graphics/font features can still require a full PDF renderer and are intentionally not claimed as fully supported.

### Layout

```text
OfficeViewer.Avalonia.NuGet/
  src/OfficeViewer.Avalonia/       Package source / 包源码
  samples/PackageConsumer/         PackageReference-only consumer / 独立消费项目
  scripts/pack-and-verify.ps1      Package verification / 打包验证
  artifacts/packages/              Generated .nupkg and .snupkg files / 生成包
```
