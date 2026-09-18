using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XlsxDocxGenerator.Services.Receipts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using SaveOptions = System.Xml.Linq.SaveOptions;

namespace XlsxDocxGenerator.Tests;

public sealed class ReceiptCutLineTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace Wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    // Phase 3 af45e59 document.xml, removing only the two existing AlternateContent
    // drawing payloads at body indexes 4 and 9. All text and flow properties remain.
    private const string BaselineFlowHash = "9D3A54015F3965CFD4B65D2ADA3C89846C14E553C2AF5867E6FCD74773A78D31";
    private const int CandidateStep = 90_000; // 2.5 mm, not a paragraph-spacing change.

    [Fact]
    public void SeparatorParagraphBordersAreAbsentAndNoParagraphWasAdded()
    {
        var document = ReadDocument(TemplatePath());
        var children = Body(document).Elements().ToArray();
        Assert.Equal(15, children.Length);
        Assert.Equal(11, Body(document).Elements(W + "p").Count());
        foreach (var index in new[] { 4, 9 })
        {
            Assert.Null(children[index].Element(W + "pPr")!.Element(W + "pBdr"));
            Assert.Empty(children[index].Descendants(W + "t"));
        }
        Assert.DoesNotContain(document.Descendants(W + "bottom"),
            border => (string?)border.Attribute(W + "val") == "dashed");
    }

    [Fact]
    public void BaselineParagraphSpacingAndParagraphMarkFormattingAreRestored()
    {
        var children = Body(ReadDocument(TemplatePath())).Elements().ToArray();
        foreach (var index in new[] { 4, 9, 10 })
        {
            var spacing = children[index].Element(W + "pPr")!.Element(W + "spacing")!;
            Assert.Equal("240", (string?)spacing.Attribute(W + "before"));
            Assert.Equal("auto", (string?)spacing.Attribute(W + "lineRule"));
            Assert.Null(spacing.Attribute(W + "after"));
            Assert.Null(spacing.Attribute(W + "line"));
        }
        Assert.Empty(children[4].Element(W + "pPr")!.Element(W + "rPr")!.Elements());
        var mark = children[9].Element(W + "pPr")!.Element(W + "rPr")!;
        Assert.Equal("28", (string?)mark.Element(W + "sz")!.Attribute(W + "val"));
        Assert.Equal("1", (string?)mark.Element(W + "b")!.Attribute(W + "val"));
        Assert.Equal("single", (string?)mark.Element(W + "u")!.Attribute(W + "val"));
    }

    [Fact]
    public void ExactlyTwoFloatingLinesOccupyTheOriginalInterReceiptParagraphs()
    {
        AssertFloatingLines(ReadDocument(TemplatePath()), 0);
    }

    [Fact]
    public void LinesHavePrintableWidthAndConsistentHorizontalShapeCoordinates()
    {
        var document = ReadDocument(TemplatePath());
        var section = Body(document).Element(W + "sectPr")!;
        var page = section.Element(W + "pgSz")!;
        var margin = section.Element(W + "pgMar")!;
        var width = ((long)page.Attribute(W + "w")! - (long)margin.Attribute(W + "left")!
            - (long)margin.Attribute(W + "right")!) * 635;
        Assert.Equal(6_840_220L, width);
        foreach (var anchor in document.Descendants(Wp + "anchor"))
        {
            Assert.Equal(width, (long)anchor.Element(Wp + "extent")!.Attribute("cx")!);
            Assert.Equal("0", (string?)anchor.Element(Wp + "extent")!.Attribute("cy"));
            var transform = anchor.Descendants(A + "xfrm").Single();
            Assert.Equal("0", (string?)transform.Element(A + "off")!.Attribute("x"));
            Assert.Equal("0", (string?)transform.Element(A + "off")!.Attribute("y"));
            Assert.Equal(width, (long)transform.Element(A + "ext")!.Attribute("cx")!);
            Assert.Equal("0", (string?)transform.Element(A + "ext")!.Attribute("cy"));
        }
    }

    [Fact]
    public void AllThreeReceiptFlowStructuresMatchPhase3ExceptTheDrawingPayloads()
    {
        var document = ReadDocument(TemplatePath());
        RemoveCutDrawingPayloads(document);
        Assert.Equal(BaselineFlowHash, Hash(Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting))));
        var tables = document.Descendants(W + "tbl").ToArray();
        Assert.Equal(3, tables.Length);
        Assert.All(tables, table =>
        {
            Assert.Equal(5, table.Elements(W + "tr").Count());
            Assert.Equal(17, table.Element(W + "tblGrid")!.Elements(W + "gridCol").Count());
        });
    }

    [Fact]
    public void ReferenceFilesAndAllOtherTemplatePartsRemainUnchanged()
    {
        Assert.Equal("3F4D65C73128B12B74A1D80CF2E65819339AAA48C729D9AE11855D30740BC38F",
            Hash(File.ReadAllBytes(Path.Combine(Root(), "reference/word.docx"))));
        Assert.Equal("6749AE8A222D9411A8978280894DCCFCC5339E5CA954BEA4C799CBE2FC4A721F",
            Hash(File.ReadAllBytes(Path.Combine(Root(), "reference/excel.xlsx"))));
        using var package = ZipFile.OpenRead(TemplatePath());
        var parts = package.Entries.Where(entry => entry.FullName != "word/document.xml")
            .OrderBy(entry => entry.FullName, StringComparer.Ordinal).Select(entry =>
            {
                using var stream = entry.Open();
                return entry.FullName + ":" + Convert.ToHexString(SHA256.HashData(stream));
            });
        Assert.Equal("8D9A3A5736A3BC85C946CEC438C4480148C9F01C32073E01EC6AB05CFE28DEB2",
            Hash(Encoding.UTF8.GetBytes(string.Join("\n", parts))));
    }

    [Fact]
    public void FloatingLinesPreservePlaceholderCountsAndFixedCellMappings()
    {
        var contract = new ReceiptTemplateContractValidator().Validate(TemplatePath());
        Assert.Equal(15, contract.OccurrenceCounts.Count);
        Assert.All(ReceiptTemplateMapping.Fields, field => Assert.Equal(3, contract.OccurrenceCounts[field.Placeholder]));
        var document = ReadDocument(TemplatePath());
        Assert.Contains("21813799", Text(document.Root!));
        foreach (var table in document.Descendants(W + "tbl"))
        {
            var rows = table.Elements(W + "tr").ToArray();
            Assert.Equal("{{PAYER}}", Text(rows[2].Elements(W + "tc").ElementAt(0)));
            Assert.Equal("{{REASON}}", Text(rows[2].Elements(W + "tc").ElementAt(9)));
            Assert.Equal("{{HANDLER}}", Text(rows[4].Elements(W + "tc").ElementAt(1)));
        }
    }

    [Fact]
    public void NormalizedFloatingDrawingsPassOffice2010SchemaValidation()
    {
        ValidateDrawings(TemplatePath());
    }

    [Fact]
    public async Task ThreeProductionGeneratedCandidatesDifferOnlyInTheirTwoLineOffsets()
    {
        var qa = Path.Combine(Root(), "artifacts/qa/phase3a2");
        var templates = Path.Combine(qa, "candidate-templates");
        Directory.CreateDirectory(templates);
        var workbookPath = Path.Combine(qa, "cut-line-input.xlsx");
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add(ReceiptWorksheetSchema.WorksheetName);
            foreach (var (header, index) in ReceiptWorksheetSchema.ExpectedHeaders.Select((value, index) => (value, index)))
                sheet.Cell(2, index + 1).Value = header;
            sheet.Cell(3, 1).Value = 112;
            sheet.Cell(3, 2).Value = 12;
            sheet.Cell(3, 3).Value = 30;
            sheet.Cell(3, 4).Value = "0051";
            sheet.Cell(3, 4).Style.NumberFormat.Format = "@";
            sheet.Cell(3, 5).Value = "A實業股份有限公司";
            sheet.Cell(3, 6).Value = 1_000;
            sheet.Cell(3, 7).Value = "運動會禮金";
            sheet.Cell(3, 8).Value = "6/6匯款";
            workbook.SaveAs(workbookPath);
        }
        var record = await new ReceiptRecordReader().ReadRecordAsync(workbookPath, ReceiptWorksheetSchema.WorksheetName, 3);
        Assert.Equal("運動會禮金", record.Reason);
        Assert.Equal("6/6匯款", record.Handler);
        var production = ReadDocument(TemplatePath());
        string? normalizedCandidate = null;
        foreach (var (name, delta) in new[] { ("A", -CandidateStep), ("B", 0), ("C", CandidateStep) })
        {
            var template = TemplatePath();
            if (delta != 0)
            {
                template = Path.Combine(templates, $"template-{name}.docx");
                File.Copy(TemplatePath(), template, overwrite: true);
                using var copy = WordprocessingDocument.Open(template, true);
                var positions = copy.MainDocumentPart!.Document.Descendants<DW.VerticalPosition>().ToArray();
                Assert.Equal(2, positions.Length);
                foreach (var position in positions)
                {
                    var offset = position.GetFirstChild<DW.PositionOffset>()!;
                    offset.Text = (int.Parse(offset.Text, CultureInfo.InvariantCulture) + delta).ToString(CultureInfo.InvariantCulture);
                }
                copy.MainDocumentPart.Document.Save();
            }
            var templateDocument = ReadDocument(template);
            AssertFloatingLines(templateDocument, delta);
            RemoveCutDrawingPayloads(templateDocument);
            var productionFlow = new XDocument(production);
            RemoveCutDrawingPayloads(productionFlow);
            Assert.Equal(Canonical(productionFlow.Root!), Canonical(templateDocument.Root!));

            var output = Path.Combine(qa, $"candidate-{name}.docx");
            File.Delete(output); // This test's ignored QA outputs only.
            await new ReceiptGenerationService(internalTemplatePath: template).GenerateReceiptAsync(record, output);
            new ReceiptGeneratedDocumentValidator().Validate(output, record, new ReceiptAmountFormatter().Format(record.Amount));
            ValidateDrawings(output);
            var generated = ReadDocument(output);
            AssertFloatingLines(generated, delta);
            Assert.Equal(LayoutFingerprint(production), LayoutFingerprint(generated));
            var text = Text(generated.Root!);
            Assert.DoesNotContain("{{", text);
            Assert.DoesNotContain("}}", text);
            Assert.DoesNotContain(record.Handler, text);
            Assert.Contains("21813799", text);
            foreach (var expected in new[] { record.Payer, record.Reason, "1120051", "壹仟元整", "中華民國112年12月30日" })
                Assert.Equal(3, text.Split(expected, StringSplitOptions.None).Length - 1);
            foreach (var table in generated.Descendants(W + "tbl"))
            {
                var rows = table.Elements(W + "tr").ToArray();
                var cells = rows[2].Elements(W + "tc").ToArray();
                Assert.Equal(record.Payer, Text(cells[0]));
                Assert.Equal(record.Reason, Text(cells[9]));
                Assert.Equal(new[] { "", "", "$", "1", "0", "0", "0" }, cells.Skip(2).Take(7).Select(Text));
                Assert.Equal(string.Empty, Text(rows[4].Elements(W + "tc").ElementAt(1)));
            }
            foreach (var offset in generated.Descendants(Wp + "positionV").Elements(Wp + "posOffset"))
                offset.Value = "0";
            var normalized = Canonical(generated.Root!);
            if (normalizedCandidate is not null) Assert.Equal(normalizedCandidate, normalized);
            normalizedCandidate = normalized;
        }
    }

    private static void AssertFloatingLines(XDocument document, int delta)
    {
        var children = Body(document).Elements().ToArray();
        Assert.Equal(15, children.Length);
        var anchors = document.Descendants(Wp + "anchor").ToArray();
        Assert.Equal(2, anchors.Length);
        Assert.Equal(2, document.Descendants(Wps + "wsp").Count());
        Assert.Empty(document.Descendants(Wp + "inline"));
        Assert.Empty(document.Descendants(Mc + "Fallback"));
        Assert.Empty(document.Descendants(A + "blip"));
        var labels = new[] { "【第三聯】", "【第二聯】", "【第一聯】" };
        for (var i = 0; i < 2; i++)
        {
            var index = 4 + i * 5;
            Assert.Same(children[index], anchors[i].Ancestors(W + "p").Single());
            Assert.Same(Body(document), children[index].Parent);
            Assert.Empty(children[index].Descendants(W + "t"));
            Assert.Null(children[index].Element(W + "pPr")!.Element(W + "pBdr"));
            Assert.Contains(labels[i], Text(children[index - 4]));
            Assert.Equal(W + "tbl", children[index - 2].Name);
            Assert.StartsWith("說明：", Text(children[index - 1]));
            Assert.Contains(labels[i + 1], Text(children[index + 1]));
            Assert.Single(anchors[i].Elements(Wp + "wrapNone"));
            Assert.Single(anchors[i].Elements().Where(element => element.Name.LocalName.StartsWith("wrap", StringComparison.Ordinal)));
            Assert.Equal("1", (string?)anchors[i].Attribute("allowOverlap"));
            Assert.Equal("0", (string?)anchors[i].Attribute("simplePos"));
            Assert.Equal("0", (string?)anchors[i].Attribute("hidden"));
            Assert.Equal("0", (string?)anchors[i].Attribute("behindDoc"));
            Assert.Equal("column", (string?)anchors[i].Element(Wp + "positionH")!.Attribute("relativeFrom"));
            Assert.Equal("0", anchors[i].Element(Wp + "positionH")!.Element(Wp + "posOffset")!.Value);
            var vertical = anchors[i].Element(Wp + "positionV")!;
            Assert.Equal("paragraph", (string?)vertical.Attribute("relativeFrom"));
            Assert.Equal((i == 0 ? 119_380 : 274_955) + delta, (int)vertical.Element(Wp + "posOffset")!);
            Assert.Equal("straightConnector1", (string?)anchors[i].Descendants(A + "prstGeom").Single().Attribute("prst"));
            var line = anchors[i].Descendants(A + "ln").Single();
            Assert.Equal("9525", (string?)line.Attribute("w")); // 0.75 pt in EMU.
            Assert.Equal("dash", (string?)line.Element(A + "prstDash")!.Attribute("val"));
            Assert.Equal("808080", (string?)line.Element(A + "solidFill")!.Element(A + "srgbClr")!.Attribute("val"));
        }
        Assert.Empty(children.Skip(10).SelectMany(element => element.Descendants(Wp + "anchor")));
    }

    private static void ValidateDrawings(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        var drawings = document.MainDocumentPart!.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().ToArray();
        Assert.Equal(2, drawings.Length);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2010);
        foreach (var drawing in drawings)
        {
            var errors = validator.Validate(drawing).Select(error => error.Description).ToArray();
            Assert.True(errors.Length == 0, string.Join("\n", errors));
        }
    }

    private static void RemoveCutDrawingPayloads(XDocument document)
    {
        var children = Body(document).Elements().ToArray();
        foreach (var index in new[] { 4, 9 })
            Assert.Single(children[index].Descendants(Mc + "AlternateContent")).Remove();
    }

    private static string LayoutFingerprint(XDocument document)
    {
        var names = new[] { "sectPr", "tblPr", "tblGrid", "trPr", "tcPr", "pPr", "rPr" }.Select(name => W + name).ToHashSet();
        return string.Join("\n", document.Descendants().Where(element => names.Contains(element.Name)).Select(Canonical));
    }

    private static string Canonical(XElement element) => element.Name + "[" +
        string.Join("|", element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)
            .OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal)
            .Select(attribute => attribute.Name + "=" + attribute.Value)) + "](" +
        string.Concat(element.Nodes().Select(node => node is XElement child ? Canonical(child) : node.ToString())) + ")";

    private static XDocument ReadDocument(string path)
    {
        using var package = ZipFile.OpenRead(path);
        using var stream = package.GetEntry("word/document.xml")!.Open();
        return XDocument.Load(stream);
    }

    private static XElement Body(XDocument document) => document.Root!.Element(W + "body")!;
    private static string Text(XElement element) => string.Concat(element.Descendants(W + "t").Select(text => text.Value));
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));
    private static string TemplatePath() => Path.Combine(Root(), ReceiptTemplateMapping.RelativeTemplatePath);
    private static string Root()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
            if (File.Exists(Path.Combine(directory.FullName, "ReceiptXcel.sln"))) return directory.FullName;
        throw new DirectoryNotFoundException("ReceiptXcel solution root was not found.");
    }
}
