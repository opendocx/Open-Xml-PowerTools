# OpenXmlPowerTools

[![.NET build and test](https://github.com/opendocx/Open-Xml-PowerTools/actions/workflows/dotnet.yml/badge.svg)](https://github.com/opendocx/Open-Xml-PowerTools/actions/workflows/dotnet.yml)

## Focus of this fork

- DocumentAssembler (Populating content in template DOCX files with data from XML)
- DocumentBuilder (Combining multiple DOCX files into a single file)
- DocumentComposer (this fork's primary contribution: integrates/wraps DocumentAssembler AND DocumentBuilder)
- focus on fixing broad range of bugs as flushed out by very diverse DOCX content
- Linux, Windows and macOS support as added by upstream (Codeuctivity)

Note on naming: this repository name still uses the historical hyphenated form (Open-Xml-PowerTools), while the
Codeuctivity fork and the .NET package/namespace use OpenXmlPowerTools. This fork is expected to drop the hyphens
in the repo name in a future rename.

See [docs/index.md](docs/index.md) for details.

## Quick start

Use `DocumentComposer` when you want DocumentAssembler-style templates *and* document insertion/concatenation in one pipeline:

```csharp
using OpenXmlPowerTools;
using System.Xml.Linq;

var template = new WmlDocument("path/to/ParentTemplate.docx");
var data = XElement.Load("path/to/data.xml");

var result = await DocumentComposer.ComposeDocument(template, data);
result.SaveAs("path/to/output.docx");
```

## When to use what

- `DocumentAssembler`: You have a Word-editable DOCX template containing placeholders, and you want to populate it from XML.
- `DocumentBuilder`: You want programmatic composition (insert/merge/concatenate) of multiple DOCX files.
- `DocumentComposer`: You want to combine DocumentAssembler + DocumentBuilder, i.e. when templates need to dynamically insert additional DOCX content.

### Other features

- Splitting DOCX/PPTX files into multiple files.
- Conversion of DOCX to HTML/CSS.
- Conversion of HTML/CSS to DOCX.
- Searching and replacing content in DOCX/PPTX using regular expressions.
- Managing tracked-revisions, including detecting tracked revisions, and accepting tracked revisions.
- Updating Charts in DOCX/PPTX files, including updating cached data, as well as the embedded XLSX.
- Retrieving metrics from DOCX files, including the hierarchy of styles used, the languages used, and the fonts used.
- Writing XLSX files using far simpler code than directly writing the markup, including a streaming approach that
  enables writing XLSX files with millions of rows.
- Extracting data (along with formatting) from spreadsheets.

## SkiaSharp migration

Earlier releases used the ImageSharp library for tasks such as decoding images in the WordprocessingML to HTML converter and validating rendering output in tests. ImageSharp's powerful API came with a more restrictive licensing model that could require commercial agreements, which proved limiting for downstream projects.

The project now uses SkiaSharp to handle these responsibilities. SkiaSharp, distributed under the permissive MIT license, provides cross-platform bindings to the Skia graphics engine. By leveraging SkiaSharp's `SKCodec` and `SKImage` APIs for image transformation and `SKColor` for color parsing, the codebase avoids licensing friction while retaining rich imaging capabilities.

## Development

- Run `dotnet build OpenXmlPowerTools.sln`
