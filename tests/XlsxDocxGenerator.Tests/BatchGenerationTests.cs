using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Errors;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Generation;
using XlsxDocxGenerator.Services.Word;

namespace XlsxDocxGenerator.Tests;

public sealed class BatchGenerationTests
{
    private readonly SelectionRuleParser _parser = new();
    private readonly RowResolver _resolver = new();

    [Fact]
    public void ParsesSingleExcelRow()
    {
        var rule = _parser.ParseExcelRows("5");

        Assert.Equal([5], Assert.IsType<SelectionRule.ExcelRows>(rule).RowNumbers);
    }

    [Fact]
    public void ParsesCommaSeparatedRows()
    {
        var rule = _parser.ParseExcelRows("5,8,10");

        Assert.Equal([5, 8, 10], Assert.IsType<SelectionRule.ExcelRows>(rule).RowNumbers);
    }

    [Fact]
    public void ParsesRange()
    {
        var rule = _parser.ParseExcelRows("10-12");

        Assert.Equal([10, 11, 12], Assert.IsType<SelectionRule.ExcelRows>(rule).RowNumbers);
    }

    [Fact]
    public void ParsesMixedRowsAndRanges()
    {
        var rule = _parser.ParseExcelRows("5-8,12,20-22");

        Assert.Equal([5, 6, 7, 8, 12, 20, 21, 22], Assert.IsType<SelectionRule.ExcelRows>(rule).RowNumbers);
    }

    [Fact]
    public void TrimsAndDeduplicatesRowsInAscendingOrder()
    {
        var rule = _parser.ParseExcelRows(" 8, 5,5,7-9 ");

        Assert.Equal([5, 7, 8, 9], Assert.IsType<SelectionRule.ExcelRows>(rule).RowNumbers);
    }

    [Theory]
    [InlineData("10-5")]
    [InlineData("0")]
    [InlineData("-1")]
    [InlineData("5-")]
    [InlineData("five")]
    [InlineData("5,,8")]
    public void RejectsMalformedRowExpressions(string expression)
    {
        var exception = Assert.Throws<SelectionValidationException>(() => _parser.ParseExcelRows(expression));

        Assert.Equal(SelectionErrorCode.InvalidSyntax, exception.Code);
    }

    [Fact]
    public void RejectsHeaderRowSelection()
    {
        var result = _resolver.Resolve(
            CreateWorksheet((2, "Alice", "A")),
            CreateTemplate(markerRow: 6),
            new SelectionRule.ExcelRows([1, 2]));

        Assert.False(result.Items[0].IsValid);
        Assert.True(result.Items[1].IsValid);
    }

    [Fact]
    public void RejectsMappingMarkerRowSelection()
    {
        var result = _resolver.Resolve(
            CreateWorksheet((2, "Alice", "A"), (6, "{{NAME}}", "")),
            CreateTemplate(markerRow: 6),
            new SelectionRule.ExcelRows([6]));

        Assert.False(result.Items.Single().IsValid);
        Assert.Equal(GenerationErrorCode.InvalidExcelDataRow, result.Items.Single().ErrorCode);
    }

    [Fact]
    public void OmitsBlankRowsFromAvailableRecords()
    {
        var result = _resolver.Resolve(
            CreateWorksheet((2, "Alice", "A"), (5, "Carol", "A")),
            CreateTemplate(markerRow: 6),
            new SelectionRule.ExcelRows([2, 3, 5]));

        Assert.Equal([2, 3, 5], result.Items.Select(item => item.RowNumber));
        Assert.False(result.Items[1].IsValid);
    }

    [Fact]
    public void LatestOneReturnsNewestFormalRecord()
    {
        var result = _resolver.Resolve(
            CreateWorksheet((2, "Alice", "A"), (3, "Bob", "B"), (5, "Carol", "A"), (6, "{{NAME}}", "")),
            CreateTemplate(markerRow: 6),
            new SelectionRule.Latest(1));

        Assert.Equal([5], result.Records.Select(record => record.RowNumber));
    }

    [Fact]
    public void LatestNReturnsAscendingExcelOrder()
    {
        var result = _resolver.Resolve(
            CreateWorksheet((2, "Alice", "A"), (3, "Bob", "B"), (5, "Carol", "A")),
            CreateTemplate(markerRow: 6),
            new SelectionRule.Latest(2));

        Assert.Equal([3, 5], result.Records.Select(record => record.RowNumber));
    }

    [Fact]
    public void LatestSkipsBlankRowsAndMarkerRow()
    {
        var result = _resolver.Resolve(
            CreateWorksheet((2, "Alice", "A"), (3, "Bob", "B"), (5, "Carol", "A"), (6, "{{NAME}}", "")),
            CreateTemplate(markerRow: 6),
            new SelectionRule.Latest(2));

        Assert.Equal([3, 5], result.Records.Select(record => record.RowNumber));
    }

    [Fact]
    public void LatestExceedingAvailableReturnsAllWithNotice()
    {
        var result = _resolver.Resolve(
            CreateWorksheet((2, "Alice", "A"), (3, "Bob", "B")),
            CreateTemplate(markerRow: 6),
            new SelectionRule.Latest(20));

        Assert.Equal([2, 3], result.Records.Select(record => record.RowNumber));
        Assert.Equal("要求 20 筆，實際找到 2 筆。", result.Notice);
    }

    [Fact]
    public void RejectsInvalidLatestCount()
    {
        var exception = Assert.Throws<SelectionValidationException>(() => _parser.ParseLatest("0"));

        Assert.Equal(SelectionErrorCode.InvalidLatestCount, exception.Code);
    }

    [Fact]
    public void SelectsBySingleMarkerToken()
    {
        var result = _resolver.Resolve(
            CreateWorksheet((2, "Alice", "A"), (3, "Bob", "B"), (5, "Carol", "A,B")),
            CreateTemplate(markerRow: 6),
            _parser.ParseMarker("OUTPUT", "A"));

        Assert.Equal([2, 5], result.Records.Select(record => record.RowNumber));
    }

    [Fact]
    public void SelectsCommaSeparatedCellTokens()
    {
        var result = _resolver.Resolve(
            CreateWorksheet((2, "Alice", "A, B,部門甲")),
            CreateTemplate(markerRow: 6),
            _parser.ParseMarker("OUTPUT", "部門甲"));

        Assert.Equal([2], result.Records.Select(record => record.RowNumber));
    }

    [Fact]
    public void MarkerMatchingIsCaseInsensitive()
    {
        var result = _resolver.Resolve(
            CreateWorksheet((2, "Alice", "A")),
            CreateTemplate(markerRow: 6),
            _parser.ParseMarker("OUTPUT", "a"));

        Assert.Single(result.Records);
    }

    [Fact]
    public void MarkerMatchingDoesNotUseSubstring()
    {
        var result = _resolver.Resolve(
            CreateWorksheet((2, "Alice", "AA")),
            CreateTemplate(markerRow: 6),
            _parser.ParseMarker("OUTPUT", "A"));

        Assert.Empty(result.Records);
    }

    [Fact]
    public void MultipleQueryMarkersUseOrSemantics()
    {
        var result = _resolver.Resolve(
            CreateWorksheet((2, "Alice", "A"), (3, "Bob", "B"), (4, "Carol", "C")),
            CreateTemplate(markerRow: 6),
            _parser.ParseMarker("OUTPUT", "A,B"));

        Assert.Equal([2, 3], result.Records.Select(record => record.RowNumber));
    }

    [Fact]
    public void MarkerWithNoMatchesReturnsEmptySelection()
    {
        var result = _resolver.Resolve(
            CreateWorksheet((2, "Alice", "A")),
            CreateTemplate(markerRow: 6),
            _parser.ParseMarker("OUTPUT", "Z"));

        Assert.Empty(result.Items);
    }

    [Fact]
    public void MissingMarkerTokenIsValidationError()
    {
        var exception = Assert.Throws<SelectionValidationException>(() =>
            _parser.ParseMarker("OUTPUT", " , "));

        Assert.Equal(SelectionErrorCode.MissingMarkerToken, exception.Code);
    }

    [Fact]
    public void MissingMarkerColumnIsValidationError()
    {
        var exception = Assert.Throws<SelectionValidationException>(() =>
            _resolver.Resolve(
                CreateWorksheet((2, "Alice", "A")),
                CreateTemplate(markerRow: 6),
                _parser.ParseMarker("NOT_FOUND", "A")));

        Assert.Equal(SelectionErrorCode.MissingMarkerColumn, exception.Code);
    }

    [Fact]
    public void ResolvesMappingFilenamePattern()
    {
        var policy = new FilenamePolicy();
        var template = CreateTemplate(markerRow: 6, pattern: "{{ID}}_{{NAME}}");
        var record = CreateRecord(2, "001", "Alice");
        template = template with
        {
            FieldMappings =
            [
                new FieldMapping { PlaceholderName = "ID", ColumnIndex = 1 },
                new FieldMapping { PlaceholderName = "NAME", ColumnIndex = 2 }
            ]
        };

        Assert.Equal("001_Alice", policy.ResolveBaseName(template.OutputFileNamePattern, template, record));
    }

    [Fact]
    public void MissingFilenamePlaceholderFailsRow()
    {
        var policy = new FilenamePolicy();
        var template = CreateTemplate(markerRow: 6, pattern: "{{UNKNOWN}}");

        var exception = Assert.Throws<GenerationException>(() =>
            policy.ResolveBaseName(template.OutputFileNamePattern, template, CreateRecord(2, "Alice")));

        Assert.Equal(GenerationErrorCode.InvalidOutputFileName, exception.Code);
    }

    [Fact]
    public void SanitizesWindowsInvalidCharacters()
    {
        var policy = new FilenamePolicy();
        var template = CreateTemplate(markerRow: 6, pattern: "{{NAME}}") with
        {
            FieldMappings = [new FieldMapping { PlaceholderName = "NAME", ColumnIndex = 1 }]
        };

        Assert.Equal("A_B_C_D_E_F_G", policy.ResolveBaseName(
            template.OutputFileNamePattern,
            template,
            CreateRecord(2, "A:B/C\\D|E?F*G")));
    }

    [Fact]
    public void SanitizesTrailingSpaceAndDot()
    {
        var policy = new FilenamePolicy();
        var template = CreateTemplate(markerRow: 6, pattern: "{{NAME}}") with
        {
            FieldMappings = [new FieldMapping { PlaceholderName = "NAME", ColumnIndex = 1 }]
        };

        Assert.Equal("name", policy.ResolveBaseName(
            template.OutputFileNamePattern,
            template,
            CreateRecord(2, "name. ")));
    }

    [Fact]
    public void SanitizesReservedWindowsDeviceName()
    {
        var policy = new FilenamePolicy();
        var template = CreateTemplate(markerRow: 6, pattern: "{{NAME}}") with
        {
            FieldMappings = [new FieldMapping { PlaceholderName = "NAME", ColumnIndex = 1 }]
        };

        Assert.Equal("_CON", policy.ResolveBaseName(
            template.OutputFileNamePattern,
            template,
            CreateRecord(2, "CON")));
    }

    [Fact]
    public void UsesRowFallbackWhenPatternIsEmpty()
    {
        var policy = new FilenamePolicy();
        var template = CreateTemplate(markerRow: 6, pattern: null);

        Assert.Equal("output-row-2", policy.ResolveBaseName(null, template, CreateRecord(2, "Alice")));
    }

    [Fact]
    public void AllocatesDeterministicDuplicateSuffixes()
    {
        using var workspace = new TemporaryWorkspace();
        var allocator = new FilenameAllocator();
        File.WriteAllText(Path.Combine(workspace.Path, "001_Alice.docx"), "existing");

        var first = allocator.Allocate(workspace.Path, "001_Alice");
        var second = allocator.Allocate(workspace.Path, "001_Alice");

        Assert.EndsWith("001_Alice (2).docx", first);
        Assert.EndsWith("001_Alice (3).docx", second);
    }

    [Fact]
    public void NormalizesDocxExtension()
    {
        var policy = new FilenamePolicy();
        var template = CreateTemplate(markerRow: 6, pattern: "file.docx");

        Assert.Equal("file", policy.ResolveBaseName(template.OutputFileNamePattern, template, CreateRecord(2, "A")));
    }

    [Fact]
    public void RejectsNonDocxExtension()
    {
        var policy = new FilenamePolicy();
        var template = CreateTemplate(markerRow: 6, pattern: "file.pdf");

        var exception = Assert.Throws<GenerationException>(() =>
            policy.ResolveBaseName(template.OutputFileNamePattern, template, CreateRecord(2, "A")));

        Assert.Equal(GenerationErrorCode.InvalidOutputFileName, exception.Code);
    }

    [Fact]
    public void ResolvesBuiltInRowFilenamePlaceholder()
    {
        var policy = new FilenamePolicy();
        var template = CreateTemplate(markerRow: 6, pattern: "document-{{ROW}}");

        Assert.Equal("document-2", policy.ResolveBaseName(template.OutputFileNamePattern, template, CreateRecord(2, "A")));
    }

    [Fact]
    public async Task BatchGeneratesMultipleSuccessfulRows()
    {
        using var workspace = new TemporaryWorkspace();
        var workbook = CreateWorkbook(workspace.Path, includeMarker: true);
        var wordTemplate = CreateDocument(workspace.Path, "姓名：{{NAME}}");
        var template = CreateTemplate(wordTemplate, markerRow: 6, pattern: "{{ROW}}_{{NAME}}");

        var result = await CreateService().GenerateBatchAsync(
            workbook,
            template,
            _parser.ParseMarker("OUTPUT", "A"),
            workspace.Path);

        Assert.Equal(2, result.SuccessCount);
        Assert.Equal(0, result.FailureCount);
        Assert.Equal([2, 5], result.Results.Select(item => item.ExcelRowNumber));
        Assert.All(result.Results, item => Assert.True(File.Exists(item.OutputPath)));
    }

    [Fact]
    public async Task BatchContinuesAfterOneRowFailure()
    {
        using var workspace = new TemporaryWorkspace();
        var workbook = CreateWorkbook(workspace.Path, includeMarker: true);
        var wordTemplate = CreateDocument(workspace.Path, "姓名：{{NAME}}");
        var template = CreateTemplate(wordTemplate, markerRow: 6, pattern: "{{ROW}}_{{NAME}}");

        var result = await CreateService().GenerateBatchAsync(
            workbook,
            template,
            _parser.ParseExcelRows("2,4,5"),
            workspace.Path);

        Assert.Equal(2, result.SuccessCount);
        Assert.Equal(1, result.FailureCount);
        Assert.Equal(4, result.Results.Single(item => item.ExcelRowNumber == 4).ExcelRowNumber);
        Assert.True(File.Exists(Path.Combine(workspace.Path, "2_Alice.docx")));
        Assert.True(File.Exists(Path.Combine(workspace.Path, "5_Carol.docx")));
    }

    [Fact]
    public async Task BatchReportsCorrectProgressOrder()
    {
        using var workspace = new TemporaryWorkspace();
        var workbook = CreateWorkbook(workspace.Path, includeMarker: true);
        var wordTemplate = CreateDocument(workspace.Path, "{{NAME}}");
        var template = CreateTemplate(wordTemplate, markerRow: 6);
        var progressValues = new List<(int Completed, int Total, int Row)>();
        var progress = new InlineProgress<BatchProgress>(value =>
            progressValues.Add((value.Completed, value.Total, value.RowNumber)));

        await CreateService().GenerateBatchAsync(
            workbook,
            template,
            _parser.ParseExcelRows("2,5"),
            workspace.Path,
            progress);

        Assert.Equal([0, 1, 2], progressValues.Select(value => value.Completed));
        Assert.Equal([2, 5], progressValues.Skip(1).Select(value => value.Row));
    }

    [Fact]
    public async Task BatchLeavesOriginalTemplateUnchanged()
    {
        using var workspace = new TemporaryWorkspace();
        var workbook = CreateWorkbook(workspace.Path, includeMarker: true);
        var wordTemplate = CreateDocument(workspace.Path, "{{NAME}}");
        var template = CreateTemplate(wordTemplate, markerRow: 6);

        await CreateService().GenerateBatchAsync(
            workbook,
            template,
            _parser.ParseExcelRows("2"),
            workspace.Path);

        Assert.Equal("{{NAME}}", ReadAllText(wordTemplate));
    }

    [Fact]
    public async Task NoMatchReturnsZeroResultWithoutGenerating()
    {
        using var workspace = new TemporaryWorkspace();
        var workbook = CreateWorkbook(workspace.Path, includeMarker: true);
        var wordTemplate = CreateDocument(workspace.Path, "{{NAME}}");
        var template = CreateTemplate(wordTemplate, markerRow: 6);

        var result = await CreateService().GenerateBatchAsync(
            workbook,
            template,
            _parser.ParseMarker("OUTPUT", "Z"),
            workspace.Path);

        Assert.Equal(0, result.TotalSelected);
        Assert.Equal(0, result.SuccessCount);
        Assert.Contains("沒有符合條件", result.Notice);
        Assert.Single(Directory.GetFiles(workspace.Path, "*.docx"));
    }

    [Fact]
    public async Task InvalidTemplateIsBatchFatal()
    {
        using var workspace = new TemporaryWorkspace();
        var workbook = CreateWorkbook(workspace.Path, includeMarker: true);
        var template = CreateTemplate(Path.Combine(workspace.Path, "missing.docx"), markerRow: 6);

        var exception = await Assert.ThrowsAsync<BatchGenerationException>(() =>
            CreateService().GenerateBatchAsync(
                workbook,
                template,
                _parser.ParseExcelRows("2"),
                workspace.Path));

        Assert.Equal(BatchFatalErrorCode.WordTemplateInvalid, exception.Code);
    }

    private GenerationService CreateService() =>
        new(new ExcelReader(), new DocxGenerator());

    private static ExcelWorksheetData CreateWorksheet(
        params (int RowNumber, string Name, string Output)[] rows)
    {
        return new ExcelWorksheetData
        {
            WorksheetName = "資料",
            HeaderRowNumber = 1,
            FirstUsedRowNumber = 1,
            LastUsedRowNumber = rows.Max(row => row.RowNumber),
            Headers = ["NAME", "OUTPUT"],
            Rows = rows.Select(row => new ExcelRowData
            {
                RowNumber = row.RowNumber,
                Cells = new Dictionary<int, string?>
                {
                    [1] = row.Name,
                    [2] = row.Output
                }
            }).ToArray()
        };
    }

    private static TemplateDefinition CreateTemplate(
        int markerRow,
        string? pattern = null) => new()
    {
        TemplateName = "測試模板",
        WordTemplatePath = "template.docx",
        WorksheetName = "資料",
        HeaderRowNumber = 1,
        MarkerRowNumber = markerRow,
        OutputFileNamePattern = pattern,
        FieldMappings = [new FieldMapping { PlaceholderName = "NAME", ColumnIndex = 1 }]
    };

    private static TemplateDefinition CreateTemplate(
        string wordTemplatePath,
        int markerRow,
        string? pattern = null) => CreateTemplate(markerRow, pattern) with
        {
            WordTemplatePath = wordTemplatePath
        };

    private static ExcelRecord CreateRecord(int rowNumber, params string?[] values) => new()
    {
        RowNumber = rowNumber,
        Values = values
            .Select((value, index) => new { Header = $"FIELD_{index + 1}", Value = value })
            .ToDictionary(item => item.Header, item => item.Value),
        ColumnValues = values
            .Select((value, index) => new { Column = index + 1, Value = value })
            .ToDictionary(item => item.Column, item => item.Value)
    };

    private static string CreateWorkbook(string directory, bool includeMarker)
    {
        var path = Path.Combine(directory, $"batch-{Guid.NewGuid():N}.xlsx");
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("資料");
        sheet.Cell(1, 1).Value = "NAME";
        sheet.Cell(1, 2).Value = "OUTPUT";
        sheet.Cell(2, 1).Value = "Alice";
        sheet.Cell(2, 2).Value = "A";
        sheet.Cell(3, 1).Value = "Bob";
        sheet.Cell(3, 2).Value = "B";
        sheet.Cell(5, 1).Value = "Carol";
        sheet.Cell(5, 2).Value = "A";
        if (includeMarker)
        {
            sheet.Cell(6, 1).Value = "{{NAME}}";
        }

        workbook.SaveAs(path);
        return path;
    }

    private static string CreateDocument(string directory, string text)
    {
        var path = Path.Combine(directory, $"template-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text(text)))));
        mainPart.Document.Save();
        return path;
    }

    private static string ReadAllText(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        return string.Concat(document.MainDocumentPart!.Document.Body!.Descendants<Text>()
            .Select(text => text.Text));
    }

    private sealed class InlineProgress<T>(Action<T> action) : IProgress<T>
    {
        public void Report(T value) => action(value);
    }

    private sealed class TemporaryWorkspace : IDisposable
    {
        public TemporaryWorkspace()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"xlsx-docx-phase3-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path))
            {
                Directory.Delete(Path, recursive: true);
            }
        }
    }
}
