using ClosedXML.Excel;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Errors;
using XlsxDocxGenerator.Services.Templates;

namespace XlsxDocxGenerator.Tests;

public sealed class TemplateServiceTests
{
    [Fact]
    public async Task ParsesNormalMarkerRow()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 1, "編號", "姓名", "地址");
            SetRow(sheet, 2, "001", "王小明", "台北市");
            SetRow(sheet, 3, "{{ID}}", "{{NAME}}", "{{ADDRESS}}");
        });

        try
        {
            var definition = await CreateService().CreateDefinitionAsync(
                path, "測試模板", "template.docx", "資料");

            Assert.Equal(3, definition.MarkerRowNumber);
            Assert.Equal(["ID", "NAME", "ADDRESS"], definition.FieldMappings.Select(x => x.PlaceholderName));
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task ParsesMultipleMarkersInOneRow()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 1, "A", "B", "C", "D");
            SetRow(sheet, 2, "{{ONE}}", "{{TWO}}", "普通文字", "{{THREE}}");
        });

        try
        {
            var definition = await CreateService().CreateDefinitionAsync(
                path, "測試模板", "", "資料");

            Assert.Equal(["ONE", "TWO", "THREE"], definition.FieldMappings.Select(x => x.PlaceholderName));
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task AllowsBlankColumnBetweenMappings()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 1, "編號", "", "地址");
            SetRow(sheet, 2, "{{ID}}", "", "{{ADDRESS}}");
        });

        try
        {
            var definition = await CreateService().CreateDefinitionAsync(
                path, "測試模板", "", "資料");

            Assert.Equal([1, 3], definition.FieldMappings.Select(x => x.ColumnIndex));
            Assert.Equal(["編號", "地址"], definition.FieldMappings.Select(x => x.HeaderName));
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task AllowsBlankHeaderName()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 1, "", "姓名");
            SetRow(sheet, 2, "{{ID}}", "{{NAME}}");
        });

        try
        {
            var definition = await CreateService().CreateDefinitionAsync(
                path, "測試模板", "", "資料");

            Assert.Null(definition.FieldMappings[0].HeaderName);
            Assert.Equal("A", definition.FieldMappings[0].ColumnLetter);
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Theory]
    [InlineData("{{}}")]
    [InlineData("{{123}}")]
    [InlineData("{{NAME ADDRESS}}")]
    [InlineData("{NAME}")]
    [InlineData("@NAME")]
    public void RejectsInvalidPlaceholderGrammar(string value)
    {
        var parser = new MarkerParser();

        Assert.False(parser.TryParse(value, out _));
    }

    [Fact]
    public async Task RejectsDuplicatePlaceholderCaseInsensitively()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 1, "A", "B");
            SetRow(sheet, 2, "{{NAME}}", "{{name}}");
        });

        try
        {
            var exception = await Assert.ThrowsAsync<TemplateDefinitionException>(() =>
                CreateService().CreateDefinitionAsync(path, "測試模板", "", "資料"));

            Assert.Equal(TemplateDefinitionErrorCode.DuplicatePlaceholder, exception.Code);
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task AutomaticallyFindsBottomMarkerRow()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 1, "A");
            SetRow(sheet, 2, "資料");
            SetRow(sheet, 5, "{{ID}}");
            sheet.Row(20).Style.Font.Bold = true;
        });

        try
        {
            var definition = await CreateService().CreateDefinitionAsync(
                path, "測試模板", "", "資料");

            Assert.Equal(5, definition.MarkerRowNumber);
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task IgnoresBlankRowsBelowMarkerRow()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 1, "A");
            SetRow(sheet, 2, "資料");
            SetRow(sheet, 5, "{{ID}}");
        });

        try
        {
            var definition = await CreateService().CreateDefinitionAsync(
                path, "測試模板", "", "資料");

            Assert.Equal(5, definition.MarkerRowNumber);
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task ReportsNoMarkerRow()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 1, "A");
            SetRow(sheet, 2, "普通資料");
        });

        try
        {
            var exception = await Assert.ThrowsAsync<TemplateDefinitionException>(() =>
                CreateService().CreateDefinitionAsync(path, "測試模板", "", "資料"));

            Assert.Equal(TemplateDefinitionErrorCode.NoMarkerRow, exception.Code);
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task ReportsAmbiguousMarkerRows()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 1, "A");
            SetRow(sheet, 2, "{{FIRST}}");
            SetRow(sheet, 4, "{{SECOND}}");
        });

        try
        {
            var exception = await Assert.ThrowsAsync<TemplateDefinitionException>(() =>
                CreateService().CreateDefinitionAsync(path, "測試模板", "", "資料"));

            Assert.Equal(TemplateDefinitionErrorCode.AmbiguousMarkerRows, exception.Code);
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task UsesExplicitMarkerRow()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 1, "A");
            SetRow(sheet, 2, "{{FIRST}}");
            SetRow(sheet, 4, "{{SECOND}}");
        });

        try
        {
            var definition = await CreateService().CreateDefinitionAsync(
                path, "測試模板", "", "資料", markerRowNumber: 4);

            Assert.Equal(4, definition.MarkerRowNumber);
            Assert.Equal("SECOND", definition.FieldMappings[0].PlaceholderName);
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task ReportsMissingWorksheet()
    {
        var path = CreateWorkbook((sheet, _) => SetHeaders(sheet, 1, "A"));

        try
        {
            var exception = await Assert.ThrowsAsync<TemplateDefinitionException>(() =>
                CreateService().CreateDefinitionAsync(path, "測試模板", "", "不存在"));

            Assert.Equal(TemplateDefinitionErrorCode.WorksheetNotFound, exception.Code);
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task SupportsHeaderRowOtherThanOne()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetRow(sheet, 3, "編號", "姓名");
            SetRow(sheet, 4, "001", "王小明");
            SetRow(sheet, 6, "{{ID}}", "{{NAME}}");
        });

        try
        {
            var definition = await CreateService().CreateDefinitionAsync(
                path, "測試模板", "", "資料", headerRowNumber: 3);

            Assert.Equal(3, definition.HeaderRowNumber);
            Assert.Equal("姓名", definition.FieldMappings[1].HeaderName);
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task MapsColumnIndexLetterAndHeaderTogether()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 1, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA");
            SetRow(sheet, 2, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "{{FIELD}}");
        });

        try
        {
            var definition = await CreateService().CreateDefinitionAsync(
                path, "測試模板", "", "資料");

            var mapping = Assert.Single(definition.FieldMappings);
            Assert.Equal(27, mapping.ColumnIndex);
            Assert.Equal("AA", mapping.ColumnLetter);
            Assert.Equal("AA", mapping.HeaderName);
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task CanonicalizesPlaceholderNamesToUppercase()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 1, "姓名");
            SetRow(sheet, 2, "{{student_name}}");
        });

        try
        {
            var definition = await CreateService().CreateDefinitionAsync(
                path, "測試模板", "", "資料");

            Assert.Equal("STUDENT_NAME", definition.FieldMappings[0].PlaceholderName);
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task ReportsInvalidExplicitMarkerRow()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 1, "A");
            SetRow(sheet, 2, "普通資料");
        });

        try
        {
            var exception = await Assert.ThrowsAsync<TemplateDefinitionException>(() =>
                CreateService().CreateDefinitionAsync(path, "測試模板", "", "資料", markerRowNumber: 2));

            Assert.Equal(TemplateDefinitionErrorCode.InvalidMarkerRow, exception.Code);
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task ReportsInvalidMarkerSyntaxInExplicitRow()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 1, "A");
            SetRow(sheet, 2, "{{123}}");
        });

        try
        {
            var exception = await Assert.ThrowsAsync<TemplateDefinitionException>(() =>
                CreateService().CreateDefinitionAsync(path, "測試模板", "", "資料", markerRowNumber: 2));

            Assert.Equal(TemplateDefinitionErrorCode.InvalidMarkerSyntax, exception.Code);
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task ReportsInvalidHeaderRow()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 1, "A");
            SetRow(sheet, 2, "{{ID}}");
        });

        try
        {
            var exception = await Assert.ThrowsAsync<TemplateDefinitionException>(() =>
                CreateService().CreateDefinitionAsync(path, "測試模板", "", "資料", headerRowNumber: 9));

            Assert.Equal(TemplateDefinitionErrorCode.InvalidHeaderRow, exception.Code);
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task ReportsMarkerRowAboveHeaderRow()
    {
        var path = CreateWorkbook((sheet, _) =>
        {
            SetHeaders(sheet, 3, "A");
            SetRow(sheet, 2, "{{ID}}");
            SetRow(sheet, 4, "資料");
        });

        try
        {
            var exception = await Assert.ThrowsAsync<TemplateDefinitionException>(() =>
                CreateService().CreateDefinitionAsync(path, "測試模板", "", "資料", headerRowNumber: 3, markerRowNumber: 2));

            Assert.Equal(TemplateDefinitionErrorCode.MarkerRowHeaderRelationshipInvalid, exception.Code);
        }
        finally
        {
            DeleteTemporaryFile(path);
        }
    }

    [Fact]
    public async Task ReportsMissingExcelFile()
    {
        var exception = await Assert.ThrowsAsync<TemplateDefinitionException>(() =>
            CreateService().CreateDefinitionAsync("missing.xlsx", "測試模板", "", "資料"));

        Assert.Equal(TemplateDefinitionErrorCode.ExcelFileNotFound, exception.Code);
    }

    private static TemplateService CreateService() =>
        new(new ExcelReader(), new MarkerParser());

    private static string CreateWorkbook(Action<IXLWorksheet, XLWorkbook> configure)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-docx-test-{Guid.NewGuid():N}.xlsx");
        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("資料");
        configure(worksheet, workbook);
        workbook.SaveAs(path);
        return path;
    }

    private static void SetHeaders(IXLWorksheet worksheet, int rowNumber, params string[] headers) =>
        SetRow(worksheet, rowNumber, headers);

    private static void SetRow(IXLWorksheet worksheet, int rowNumber, params string[] values)
    {
        for (var index = 0; index < values.Length; index++)
        {
            worksheet.Cell(rowNumber, index + 1).Value = values[index];
        }
    }

    private static void DeleteTemporaryFile(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }
}
