using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Word;

namespace XlsxDocxGenerator.Tests;

public sealed class Phase4BWordRegressionTests
{
    [Fact]
    public async Task ScannerCoversAllSupportedPartsInDeterministicOrder()
    {
        var directory = CreateDirectory();
        try
        {
            var path = Path.Combine(directory, "template.docx");
            CreateScopedDocument(
                path,
                "Body {{BODY}}",
                new PartSpec(true, HeaderFooterValues.Default, ["Header {{HEADER}}"]),
                new PartSpec(true, HeaderFooterValues.First, ["First {{FIRST}}"]),
                new PartSpec(false, HeaderFooterValues.Default, ["Footer {{FOOTER}}"]),
                new PartSpec(false, HeaderFooterValues.Even, ["Even {{EVEN}}"]));

            var reader = new DocxTemplateReader();
            var first = await reader.ScanPlaceholdersAsync(path);
            var second = await reader.ScanPlaceholdersAsync(path);

            Assert.Equal(first, second);
            Assert.Equal(["BODY", "HEADER", "FIRST", "FOOTER", "EVEN"], first);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public async Task ScannerIgnoresMalformedHeaderAndFooterPlaceholders()
    {
        var directory = CreateDirectory();
        try
        {
            var path = Path.Combine(directory, "template.docx");
            CreateScopedDocument(
                path,
                "Body {{GOOD}}",
                new PartSpec(true, HeaderFooterValues.Default, ["{{BAD FIELD}} {{123}}"]),
                new PartSpec(false, HeaderFooterValues.Default, ["{BAD} {{GOOD_FOOTER}}"]));

            var names = await new DocxTemplateReader().ScanPlaceholdersAsync(path);

            Assert.Equal(["GOOD", "GOOD_FOOTER"], names);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public async Task ScannerDeduplicatesCanonicalNamesAcrossParts()
    {
        var directory = CreateDirectory();
        try
        {
            var path = Path.Combine(directory, "template.docx");
            CreateScopedDocument(
                path,
                "Body {{name}}",
                new PartSpec(true, HeaderFooterValues.Default, ["Header {{NAME}}"]),
                new PartSpec(false, HeaderFooterValues.Default, ["Footer {{Name}}"]));

            var names = await new DocxTemplateReader().ScanPlaceholdersAsync(path);

            Assert.Equal(["NAME"], names);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public async Task ReplacesHeaderFooterTablesAndDifferentParts()
    {
        var directory = CreateDirectory();
        try
        {
            var path = Path.Combine(directory, "template.docx");
            CreateScopedDocument(
                path,
                "Body {{BODY}}",
                new PartSpec(true, HeaderFooterValues.Default, ["Table {{HEADER}}"], true),
                new PartSpec(false, HeaderFooterValues.Default, ["Table {{FOOTER}}"], true));
            var output = Path.Combine(directory, "output.docx");
            var template = Template(path, ["BODY", "HEADER", "FOOTER"]);

            await new DocxGenerator().GenerateAsync(
                template,
                Record(("BODY", "body"), ("HEADER", "header"), ("FOOTER", "footer")),
                output);

            var text = ReadAllParts(output);
            Assert.Contains("Body body", text);
            Assert.Contains("Table header", text);
            Assert.Contains("Table footer", text);
            using var reopened = WordprocessingDocument.Open(output, false);
            Assert.NotNull(reopened.MainDocumentPart?.Document.Body);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public async Task ReplacesCrossRunHeaderAndFooterWithSpecialCharactersAndWhitespace()
    {
        var directory = CreateDirectory();
        try
        {
            var path = Path.Combine(directory, "template.docx");
            CreateScopedDocument(
                path,
                "Body {{EMPTY}}",
                new PartSpec(true, HeaderFooterValues.Default, ["Header {{HE", "ADER}}"]),
                new PartSpec(false, HeaderFooterValues.Default, ["Footer {{FO", "OTER}}"]));
            var output = Path.Combine(directory, "output.docx");

            await new DocxGenerator().GenerateAsync(
                Template(path, ["EMPTY", "HEADER", "FOOTER"]),
                Record(("EMPTY", null), ("HEADER", " & < > "), ("FOOTER", "footer")),
                output);

            var text = ReadAllParts(output);
            Assert.Contains("Header  & < > ", text);
            Assert.Contains("Footer footer", text);
            Assert.Contains("Body ", text);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public async Task ReplacesFourRunHeaderAndCrossRunFooter()
    {
        var directory = CreateDirectory();
        try
        {
            var path = Path.Combine(directory, "template.docx");
            CreateScopedDocument(
                path,
                "Body",
                new PartSpec(true, HeaderFooterValues.Default, ["{{", "HE", "AD", "ER", "}}"]),
                new PartSpec(false, HeaderFooterValues.Default, ["{{FO", "OTER}}"]));
            var output = Path.Combine(directory, "output.docx");

            await new DocxGenerator().GenerateAsync(
                Template(path, ["HEADER", "FOOTER"]),
                Record(("HEADER", "header"), ("FOOTER", "footer")),
                output);

            var text = ReadAllParts(output);
            Assert.Contains("header", text);
            Assert.Contains("footer", text);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public async Task SupportsFirstFooterAndEvenHeader()
    {
        var directory = CreateDirectory();
        try
        {
            var path = Path.Combine(directory, "template.docx");
            CreateScopedDocument(
                path,
                "Body",
                new PartSpec(true, HeaderFooterValues.Even, ["Even header {{EVEN_HEADER}}"]),
                new PartSpec(false, HeaderFooterValues.First, ["First footer {{FIRST_FOOTER}}"]));
            var output = Path.Combine(directory, "output.docx");

            await new DocxGenerator().GenerateAsync(
                Template(path, ["EVEN_HEADER", "FIRST_FOOTER"]),
                Record(("EVEN_HEADER", "EH"), ("FIRST_FOOTER", "FF")),
                output);

            var text = ReadAllParts(output);
            Assert.Contains("Even header EH", text);
            Assert.Contains("First footer FF", text);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public async Task SamePlaceholderAcrossBodyHeaderFooterIsReplacedOncePerPart()
    {
        var directory = CreateDirectory();
        try
        {
            var path = Path.Combine(directory, "template.docx");
            CreateScopedDocument(
                path,
                "Body {{NAME}}",
                new PartSpec(true, HeaderFooterValues.Default, ["Header {{NAME}}"]),
                new PartSpec(false, HeaderFooterValues.Default, ["Footer {{NAME}}"]));
            var output = Path.Combine(directory, "output.docx");

            await new DocxGenerator().GenerateAsync(
                Template(path, ["NAME"]),
                Record(("NAME", "Alice")),
                output);

            var text = ReadAllParts(output);
            Assert.Equal(3, Count(text, "Alice"));
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public async Task SameHeaderPartReferencedByTwoSectionsIsProcessedOnce()
    {
        var directory = CreateDirectory();
        try
        {
            var path = Path.Combine(directory, "template.docx");
            CreateSharedPartDocument(path, true);
            var output = Path.Combine(directory, "output.docx");

            await new DocxGenerator().GenerateAsync(
                Template(path, ["NAME"]),
                Record(("NAME", "Alice")),
                output);

            Assert.Equal(1, Count(ReadAllParts(output), "Alice"));
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public async Task SameFooterPartReferencedByTwoSectionsIsProcessedOnce()
    {
        var directory = CreateDirectory();
        try
        {
            var path = Path.Combine(directory, "template.docx");
            CreateSharedPartDocument(path, false);
            var output = Path.Combine(directory, "output.docx");

            await new DocxGenerator().GenerateAsync(
                Template(path, ["NAME"]),
                Record(("NAME", "Alice")),
                output);

            Assert.Equal(1, Count(ReadAllParts(output), "Alice"));
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    private static TemplateDefinition Template(string path, IReadOnlyList<string> names) => new()
    {
        TemplateName = "Word regression",
        WordTemplatePath = path,
        PreferredWorksheetName = "資料",
        HeaderRowNumber = 1,
        FieldMappings = names.Select((name, index) => new FieldMapping
        {
            PlaceholderName = name,
            ColumnIndex = index + 1
        }).ToArray()
    };

    private static ExcelRecord Record(params (string Name, string? Value)[] values) => new()
    {
        RowNumber = 2,
        ColumnValues = values.Select((item, index) => new { item.Value, Column = index + 1 })
            .ToDictionary(item => item.Column, item => item.Value),
        Values = values.ToDictionary(item => item.Name, item => item.Value)
    };

    private static void CreateScopedDocument(string path, string bodyText, params PartSpec[] parts)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var body = new Body(new Paragraph(new Run(new Text(bodyText))));
        main.Document = new Document(body);
        var section = body.AppendChild(new SectionProperties());

        foreach (var partSpec in parts)
        {
            if (partSpec.IsHeader)
            {
                var part = main.AddNewPart<HeaderPart>();
                part.Header = new Header(CreateContent(partSpec));
                part.Header.Save();
                section.Append(new HeaderReference
                {
                    Type = partSpec.Type,
                    Id = main.GetIdOfPart(part)
                });
            }
            else
            {
                var part = main.AddNewPart<FooterPart>();
                part.Footer = new Footer(CreateContent(partSpec));
                part.Footer.Save();
                section.Append(new FooterReference
                {
                    Type = partSpec.Type,
                    Id = main.GetIdOfPart(part)
                });
            }
        }

        main.Document.Save();
    }

    private static OpenXmlElement CreateContent(PartSpec spec)
    {
        var runs = spec.RunParts.Select(text => new Run(new Text(text))).ToArray();
        if (!spec.Table)
        {
            return new Paragraph(runs);
        }

        return new Table(new TableRow(new TableCell(new Paragraph(runs))));
    }

    private static void CreateSharedPartDocument(string path, bool header)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var body = new Body(new Paragraph(new Run(new Text("Body"))));
        main.Document = new Document(body);
        var section1 = body.AppendChild(new SectionProperties());
        var section2Paragraph = body.AppendChild(new Paragraph(new Run(new Text("Section 2"))));
        var section2 = body.AppendChild(new SectionProperties());

        if (header)
        {
            var part = main.AddNewPart<HeaderPart>();
            part.Header = new Header(new Paragraph(new Run(new Text("Header {{NAME}}"))));
            part.Header.Save();
            var id = main.GetIdOfPart(part);
            section1.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = id });
            section2.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = id });
        }
        else
        {
            var part = main.AddNewPart<FooterPart>();
            part.Footer = new Footer(new Paragraph(new Run(new Text("Footer {{NAME}}"))));
            part.Footer.Save();
            var id = main.GetIdOfPart(part);
            section1.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = id });
            section2.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = id });
        }

        main.Document.Save();
    }

    private static string ReadAllParts(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        var parts = new List<string> { document.MainDocumentPart!.Document.Body!.InnerText };
        parts.AddRange(document.MainDocumentPart.HeaderParts.Select(part => part.Header?.InnerText ?? string.Empty));
        parts.AddRange(document.MainDocumentPart.FooterParts.Select(part => part.Footer?.InnerText ?? string.Empty));
        return string.Join(" ", parts);
    }

    private static int Count(string source, string value) =>
        source.Split(value, StringSplitOptions.None).Length - 1;

    private static string CreateDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), $"docxcel-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return path;
    }

    private sealed record PartSpec(
        bool IsHeader,
        HeaderFooterValues Type,
        IReadOnlyList<string> RunParts,
        bool Table = false);
}
