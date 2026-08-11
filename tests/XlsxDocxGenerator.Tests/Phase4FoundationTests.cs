using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Errors;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Templates;
using XlsxDocxGenerator.Services.Word;

namespace XlsxDocxGenerator.Tests;

public sealed class Phase4FoundationTests
{
    [Fact]
    public void CreationMarkerRowDoesNotExcludeSameNumberInAnotherWorkbook()
    {
        var worksheet = Worksheet((2, "Alice"), (3, "Bob"));
        var template = Template(creationMarkerRow: 3);

        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template);
        var result = new RowResolver().Resolve(
            worksheet,
            template,
            new SelectionRule.ExcelRows([3]),
            context.RuntimeMappingMarkerRowNumber);

        Assert.Null(context.RuntimeMappingMarkerRowNumber);
        Assert.True(result.Items.Single().IsValid);
    }

    [Fact]
    public void DetectsOnlyCompleteCurrentMappingMarker()
    {
        var worksheet = WorksheetWithTwoMappings(
            (2, "Alice", "Taipei"),
            (3, "Bob", "Kaohsiung"),
            (4, "{{NAME}}", "{{ADDRESS}}"));
        var template = Template(creationMarkerRow: 2) with
        {
            FieldMappings = [Mapping("NAME", 1, "Name"), Mapping("ADDRESS", 2, "Address")]
        };

        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template);

        Assert.Equal(4, context.RuntimeMappingMarkerRowNumber);
        Assert.True(new RowResolver().Resolve(
            worksheet,
            template,
            new SelectionRule.ExcelRows([3]),
            context.RuntimeMappingMarkerRowNumber).Items.Single().IsValid);
    }

    [Fact]
    public void RuntimeMarkerIsExcludedFromExplicitSelection()
    {
        var worksheet = WorksheetWithTwoMappings(
            (2, "Alice", "Taipei"),
            (4, "{{NAME}}", "{{ADDRESS}}"));
        var template = Template(creationMarkerRow: 2) with
        {
            FieldMappings = [Mapping("NAME", 1, "Name"), Mapping("ADDRESS", 2, "Address")]
        };
        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template);

        var item = new RowResolver().Resolve(
            worksheet,
            template,
            new SelectionRule.ExcelRows([4]),
            context.RuntimeMappingMarkerRowNumber).Items.Single();

        Assert.False(item.IsValid);
    }

    [Fact]
    public void ReportsExcelSchemaMismatchWithoutRemapping()
    {
        var worksheet = Worksheet((2, "Alice"));
        var template = Template(creationMarkerRow: null) with
        {
            FieldMappings = [Mapping("NAME", 1, "Full Name")]
        };

        var exception = Assert.Throws<TemplateDefinitionException>(() =>
            new RuntimeWorksheetSchemaValidator().Validate(worksheet, template));

        Assert.Equal(TemplateDefinitionErrorCode.ExcelSchemaMismatch, exception.Code);
    }

    [Fact]
    public void SavesAndLoadsVersionedTemplateJson()
    {
        var path = Path.Combine(Path.GetTempPath(), $"docxcel-{Guid.NewGuid():N}.docxcel.json");
        try
        {
            var service = new TemplateService(new ExcelReader(), new MarkerParser());
            var source = Template(7) with { OutputFileNamePattern = "result-{{ROW}}" };
            service.SaveTemplate(path, source);
            var loaded = service.LoadTemplate(path);

            Assert.Equal(1, loaded.SchemaVersion);
            Assert.Equal(source.TemplateName, loaded.TemplateName);
            Assert.Equal(source.PreferredWorksheetName, loaded.PreferredWorksheetName);
            Assert.Equal(source.CreationMappingMarkerRowNumber, loaded.CreationMappingMarkerRowNumber);
            Assert.DoesNotContain("ExcelPath", File.ReadAllText(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void RejectsFutureTemplateVersion()
    {
        var path = Path.Combine(Path.GetTempPath(), $"docxcel-{Guid.NewGuid():N}.docxcel.json");
        try
        {
            File.WriteAllText(path, "{\"schemaVersion\":99,\"templateName\":\"T\",\"wordTemplatePath\":\"template.docx\",\"preferredWorksheetName\":\"Data\",\"headerRowNumber\":1,\"fieldMappings\":[]}");
            var exception = Assert.Throws<TemplatePersistenceException>(() =>
                new TemplateService(new ExcelReader(), new MarkerParser()).LoadTemplate(path));
            Assert.Equal(TemplatePersistenceErrorCode.UnsupportedTemplateVersion, exception.Code);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task ScansAndReplacesHeaderFooterAndBody()
    {
        var directory = CreateTempDirectory();
        try
        {
            var templatePath = Path.Combine(directory, "template.docx");
            CreateDocument(templatePath, "Body {{NAME}}", "Header {{NAME}}", "", "Footer {{NAME}}");
            var outputPath = Path.Combine(directory, "output.docx");
            var template = Template(null) with { WordTemplatePath = templatePath };
            var record = Record("Alice");

            var names = await new DocxTemplateReader().ScanPlaceholdersAsync(templatePath);
            await new DocxGenerator().GenerateAsync(template, record, outputPath);

            Assert.Equal(["NAME"], names);
            Assert.Contains("Body Alice", ReadAllParts(outputPath));
            Assert.Contains("Header Alice", ReadAllParts(outputPath));
            Assert.Contains("Footer Alice", ReadAllParts(outputPath));
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public async Task ReplacesCrossRunHeaderPlaceholder()
    {
        var directory = CreateTempDirectory();
        try
        {
            var templatePath = Path.Combine(directory, "template.docx");
            CreateDocument(templatePath, "Body", "{{NA", "ME}}", footerText: "Footer");
            var outputPath = Path.Combine(directory, "output.docx");
            await new DocxGenerator().GenerateAsync(
                Template(null) with { WordTemplatePath = templatePath },
                Record("Alice"),
                outputPath);
            Assert.Contains("Alice", ReadAllParts(outputPath));
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public async Task SupportsHeaderTableAndFirstEvenParts()
    {
        var directory = CreateTempDirectory();
        try
        {
            var templatePath = Path.Combine(directory, "template.docx");
            CreateVariantPartsDocument(templatePath);
            var outputPath = Path.Combine(directory, "output.docx");

            await new DocxGenerator().GenerateAsync(
                Template(null) with { WordTemplatePath = templatePath },
                Record("Alice"),
                outputPath);

            var text = ReadAllParts(outputPath);
            Assert.Contains("Table Alice", text);
            Assert.Contains("First Alice", text);
            Assert.Contains("Even Alice", text);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    private static TemplateDefinition Template(int? creationMarkerRow) => new()
    {
        TemplateName = "Test",
        WordTemplatePath = "template.docx",
        PreferredWorksheetName = "Data",
        HeaderRowNumber = 1,
        CreationMappingMarkerRowNumber = creationMarkerRow,
        FieldMappings = [Mapping("NAME", 1, "Name")]
    };

    private static FieldMapping Mapping(string name, int column, string? header) => new()
    {
        PlaceholderName = name,
        ColumnIndex = column,
        HeaderName = header
    };

    private static ExcelWorksheetData Worksheet(params (int Row, string Value)[] rows) => new()
    {
        WorksheetName = "Data",
        HeaderRowNumber = 1,
        FirstUsedRowNumber = 1,
        LastUsedRowNumber = rows.Max(row => row.Row),
        Headers = ["Name"],
        Rows = rows.Select(row => new ExcelRowData
        {
            RowNumber = row.Row,
            Cells = new Dictionary<int, string?> { [1] = row.Value }
        }).ToArray()
    };

    private static ExcelWorksheetData WorksheetWithTwoMappings(
        params (int Row, string Name, string Address)[] rows) => new()
    {
        WorksheetName = "Data",
        HeaderRowNumber = 1,
        FirstUsedRowNumber = 1,
        LastUsedRowNumber = rows.Max(row => row.Row),
        Headers = ["Name", "Address"],
        Rows = rows.Select(row => new ExcelRowData
        {
            RowNumber = row.Row,
            Cells = new Dictionary<int, string?>
            {
                [1] = row.Name,
                [2] = row.Address
            }.Where(pair => !string.IsNullOrEmpty(pair.Value))
             .ToDictionary(pair => pair.Key, pair => pair.Value)
        }).ToArray()
    };

    private static ExcelRecord Record(string value) => new()
    {
        RowNumber = 2,
        ColumnValues = new Dictionary<int, string?> { [1] = value },
        Values = new Dictionary<string, string?> { ["Name"] = value }
    };

    private static void CreateDocument(string path, string bodyText, string headerText, string headerSecondPart, string footerText)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var body = new Body(new Paragraph(new Run(new Text(bodyText))));
        main.Document = new Document(body);

        var header = main.AddNewPart<HeaderPart>();
        header.Header = new Header(new Paragraph(new Run(new Text(headerText)), new Run(new Text(headerSecondPart))));
        header.Header.Save();
        var footer = main.AddNewPart<FooterPart>();
        footer.Footer = new Footer(new Paragraph(new Run(new Text(footerText))));
        footer.Footer.Save();

        var sections = body.AppendChild(new SectionProperties());
        sections.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(header) });
        sections.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footer) });
        main.Document.Save();
    }

    private static void CreateVariantPartsDocument(string path)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var body = new Body(new Paragraph(new Run(new Text("Body"))));
        main.Document = new Document(body);

        var defaultHeader = main.AddNewPart<HeaderPart>();
        defaultHeader.Header = new Header(new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("Table {{NAME}}")))))));
        defaultHeader.Header.Save();
        var firstHeader = main.AddNewPart<HeaderPart>();
        firstHeader.Header = new Header(new Paragraph(new Run(new Text("First {{NAME}}"))));
        firstHeader.Header.Save();
        var evenFooter = main.AddNewPart<FooterPart>();
        evenFooter.Footer = new Footer(new Paragraph(new Run(new Text("Even {{NAME}}"))));
        evenFooter.Footer.Save();

        var sections = body.AppendChild(new SectionProperties());
        sections.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(defaultHeader) });
        sections.Append(new HeaderReference { Type = HeaderFooterValues.First, Id = main.GetIdOfPart(firstHeader) });
        sections.Append(new FooterReference { Type = HeaderFooterValues.Even, Id = main.GetIdOfPart(evenFooter) });
        main.Document.Save();
    }

    private static string ReadAllParts(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        var parts = new List<string>
        {
            document.MainDocumentPart!.Document.Body!.InnerText
        };
        parts.AddRange(document.MainDocumentPart.HeaderParts.Select(part => part.Header?.InnerText ?? string.Empty));
        parts.AddRange(document.MainDocumentPart.FooterParts.Select(part => part.Footer?.InnerText ?? string.Empty));
        return string.Join(" ", parts);
    }

    private static string CreateTempDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), $"docxcel-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return path;
    }
}
