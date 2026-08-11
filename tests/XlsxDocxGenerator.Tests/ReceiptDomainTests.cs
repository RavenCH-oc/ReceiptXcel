using ClosedXML.Excel;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Receipts;
using XlsxDocxGenerator.Services.Word;

namespace XlsxDocxGenerator.Tests;

public sealed class ReceiptDomainTests
{
    [Fact]
    public void ExactReceiptHeadersAreAccepted()
    {
        var worksheet = CreateWorksheetData(ReceiptWorksheetSchema.ExpectedHeaders);

        new ReceiptWorksheetSchema().Validate(worksheet);
    }

    [Fact]
    public void ChangedHeaderIsRejectedWithoutRemapping()
    {
        var headers = ReceiptWorksheetSchema.ExpectedHeaders.ToArray();
        headers[5] = "金額";

        var exception = Assert.Throws<ReceiptValidationException>(
            () => new ReceiptWorksheetSchema().Validate(CreateWorksheetData(headers)));

        Assert.Equal(ReceiptValidationErrorCode.ExcelSchemaMismatch, exception.Code);
        Assert.Contains("F 欄預期為「數字金額」", exception.Message);
    }

    [Fact]
    public void ReorderedHeaderIsRejectedWithoutRemapping()
    {
        var headers = ReceiptWorksheetSchema.ExpectedHeaders.ToArray();
        (headers[0], headers[1]) = (headers[1], headers[0]);

        var exception = Assert.Throws<ReceiptValidationException>(
            () => new ReceiptWorksheetSchema().Validate(CreateWorksheetData(headers)));

        Assert.Equal(ReceiptValidationErrorCode.ExcelSchemaMismatch, exception.Code);
        Assert.Contains("A 欄預期為「年」", exception.Message);
    }

    [Fact]
    public async Task ReceiptSerialKeepsExcelDisplayedLeadingZeroes()
    {
        var path = CreateWorkbook();
        try
        {
            var record = await new ReceiptRecordReader().ReadRecordAsync(path, 3);

            Assert.Equal("0051", record.ReceiptSerial);
            Assert.Equal("1120051", record.ReceiptNumber);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task NormalReceiptRecordIsReadFromFixedWorksheet()
    {
        var path = CreateWorkbook();
        try
        {
            var record = await new ReceiptRecordReader().ReadRecordAsync(path, 3);

            Assert.Equal(3, record.ExcelRowNumber);
            Assert.Equal(112, record.ROCYear);
            Assert.Equal(12, record.Month);
            Assert.Equal(30, record.Day);
            Assert.Equal("A實業股份有限公司", record.Payer);
            Assert.Equal(1000, record.Amount);
            Assert.Equal("運動會禮金", record.Reason);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("1000", 1000)]
    [InlineData("1,000", 1000)]
    [InlineData("60000", 60000)]
    [InlineData("60,000", 60000)]
    public void ReceiptAmountParserAcceptsOnlySupportedIntegerForms(string input, int expected)
    {
        Assert.Equal(expected, new ReceiptAmountParser().Parse(input));
    }

    [Theory]
    [InlineData("$1,000")]
    [InlineData("1000.00")]
    [InlineData("1,000.50")]
    [InlineData("-1000")]
    [InlineData("1,000元")]
    [InlineData("abc")]
    public void ReceiptAmountParserRejectsUnsupportedForms(string input)
    {
        var exception = Assert.Throws<ReceiptValidationException>(
            () => new ReceiptAmountParser().Parse(input));

        Assert.Equal(ReceiptValidationErrorCode.InvalidAmount, exception.Code);
    }

    [Fact]
    public void ReceiptAmountParserRejectsAmountOverFormLimitWithFriendlyRowMessage()
    {
        var exception = Assert.Throws<ReceiptValidationException>(
            () => new ReceiptAmountParser().Parse("10,000,000", 7));

        Assert.Equal(ReceiptValidationErrorCode.AmountExceedsLimit, exception.Code);
        Assert.Equal(
            "第 7 列金額 10,000,000 超過表單上限 9,999,999，請確認 Excel 資料。",
            exception.Message);
    }

    [Theory]
    [InlineData(300, "", "", "", "$", "3", "0", "0")]
    [InlineData(1000, "", "", "$", "1", "0", "0", "0")]
    [InlineData(12000, "", "$", "1", "2", "0", "0", "0")]
    [InlineData(60000, "", "$", "6", "0", "0", "0", "0")]
    [InlineData(888888, "$", "8", "8", "8", "8", "8", "8")]
    [InlineData(1000000, "1", "0", "0", "0", "0", "0", "0")]
    [InlineData(9999999, "9", "9", "9", "9", "9", "9", "9")]
    public void AmountDigitFormatterUsesSevenCellsAndPlacesDollarToTheLeft(
        int amount,
        string million,
        string hundredThousand,
        string tenThousand,
        string thousand,
        string hundred,
        string ten,
        string one)
    {
        var cells = new AmountDigitFormatter().Format(amount);

        Assert.Equal(
            new[] { million, hundredThousand, tenThousand, thousand, hundred, ten, one },
            new[]
            {
                cells.Million,
                cells.HundredThousand,
                cells.TenThousand,
                cells.Thousand,
                cells.Hundred,
                cells.Ten,
                cells.One
            });
    }

    [Theory]
    [InlineData(100, "壹佰元整")]
    [InlineData(1000, "壹仟元整")]
    [InlineData(12000, "壹萬貳仟元整")]
    [InlineData(8888, "捌仟捌佰捌拾捌元整")]
    [InlineData(101, "壹佰零壹元整")]
    [InlineData(1005, "壹仟零伍元整")]
    [InlineData(1010, "壹仟零壹拾元整")]
    [InlineData(10001, "壹萬零壹元整")]
    [InlineData(10010, "壹萬零壹拾元整")]
    [InlineData(10100, "壹萬零壹佰元整")]
    [InlineData(100001, "壹拾萬零壹元整")]
    [InlineData(1000000, "壹佰萬元整")]
    public void ChineseFinancialAmountFormatterHandlesFinancialZeroRules(
        int amount,
        string expected)
    {
        Assert.Equal(expected, new ChineseFinancialAmountFormatter().Format(amount));
    }

    [Fact]
    public void DerivedReceiptAmountIsDeterministicAndCentralized()
    {
        var formatted = new ReceiptAmountFormatter().Format(300);

        Assert.Equal(300, formatted.OriginalAmount);
        Assert.Equal("", formatted.MillionCell);
        Assert.Equal("", formatted.HundredThousandCell);
        Assert.Equal("", formatted.TenThousandCell);
        Assert.Equal("$", formatted.ThousandCell);
        Assert.Equal("3", formatted.HundredCell);
        Assert.Equal("0", formatted.TenCell);
        Assert.Equal("0", formatted.OneCell);
        Assert.Equal("參佰元整", formatted.ChineseUppercase);
    }

    [Fact]
    public void ReceiptTemplateMappingMapsHandlerToTheThreeJunctions()
    {
        var record = new ReceiptRecord
        {
            ExcelRowNumber = 3,
            ROCYear = 112,
            Month = 12,
            Day = 30,
            ReceiptSerial = "0051",
            Payer = "付款人",
            Amount = 1000,
            Reason = "事由",
            Handler = "承辦人"
        };
        var values = ReceiptTemplateMapping.CreateValues(
            record,
            new ReceiptAmountFormatter().Format(record.Amount));

        Assert.Equal("承辦人", values["HANDLER"]);
        Assert.Contains(ReceiptTemplateMapping.Fields, field =>
            field.Placeholder == "HANDLER"
            && field.TargetDescription.Contains("經手人", StringComparison.Ordinal));
    }

    [Fact]
    public async Task InternalReceiptTemplateContainsEveryMappedPlaceholder()
    {
        var templatePath = Path.Combine(
            AppContext.BaseDirectory,
            "Assets",
            "Templates",
            "receipt-template.docx");

        var placeholders = await new DocxTemplateReader().ScanPlaceholdersAsync(templatePath);

        Assert.Equal(
            ReceiptTemplateMapping.Fields
                .Select(field => field.Placeholder)
                .OrderBy(name => name),
            placeholders.OrderBy(name => name));
    }

    private static ExcelWorksheetData CreateWorksheetData(IReadOnlyList<string> headers) =>
        new()
        {
            WorksheetName = ReceiptWorksheetSchema.WorksheetName,
            HeaderRowNumber = ReceiptWorksheetSchema.HeaderRowNumber,
            FirstUsedRowNumber = 1,
            LastUsedRowNumber = 3,
            Headers = headers,
            Rows = []
        };

    private static string CreateWorkbook()
    {
        var path = Path.Combine(
            Path.GetTempPath(),
            $"receipt-xcel-test-{Guid.NewGuid():N}.xlsx");
        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet(ReceiptWorksheetSchema.WorksheetName);
        worksheet.Cell(1, 1).Value = "表單標題";
        worksheet.Range(1, 1, 1, 7).Merge();
        for (var column = 0; column < ReceiptWorksheetSchema.ExpectedHeaders.Count; column++)
        {
            worksheet.Cell(2, column + 1).Value = ReceiptWorksheetSchema.ExpectedHeaders[column];
        }

        worksheet.Cell(3, 1).Value = 112;
        worksheet.Cell(3, 2).Value = 12;
        worksheet.Cell(3, 3).Value = 30;
        worksheet.Cell(3, 4).Value = "0051";
        worksheet.Cell(3, 4).Style.NumberFormat.Format = "@";
        worksheet.Cell(3, 5).Value = "A實業股份有限公司";
        worksheet.Cell(3, 6).Value = 1000;
        worksheet.Cell(3, 6).Style.NumberFormat.Format = "#,##0";
        worksheet.Cell(3, 7).Value = "運動會禮金";
        workbook.SaveAs(path);
        return path;
    }
}
