using System;
using System.Collections.Generic;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>Failure Workbook 资源预检职责级测试。</summary>
public sealed class NpoiFailureWorkbookPreflightTest
{
    [Fact]
    public void FailureOptions_NewBudgets_ShouldRejectNonPositiveValues()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ExcelImportFailureOptions { MaxCopiedPictureBytes = 0 }.Validate());
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ExcelImportFailureOptions { MaxEstimatedTargetObjects = 0 }.Validate());
    }

    [Fact]
    public void Validate_CandidateErrorRowBudget_ShouldRejectBeforeBuilder()
    {
        using var workbook = new XSSFWorkbook();
        workbook.CreateSheet("Data");
        var errors = new[]
        {
            new ExcelImportError(ExcelImportErrorCode.InvalidInput, "invalid", "Data", 2, 1, "Value"),
            new ExcelImportError(ExcelImportErrorCode.InvalidInput, "invalid", "Data", 3, 1, "Value")
        };
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
            MaxCandidateErrorRows = 1
        };

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            NpoiFailureWorkbookPreflight.Validate(workbook, errors,
                new Dictionary<string, ExcelSheetImportRequest>(), options));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    [Fact]
    public void Validate_CandidateErrorRowBudget_ShouldAllowBoundary()
    {
        using var workbook = new XSSFWorkbook();
        workbook.CreateSheet("Data");
        var errors = new[]
        {
            new ExcelImportError(ExcelImportErrorCode.InvalidInput, "invalid", "Data", 2, 1, "Value")
        };
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
            MaxCandidateErrorRows = 1
        };

        NpoiFailureWorkbookPreflight.Validate(workbook, errors,
            new Dictionary<string, ExcelSheetImportRequest>(), options);
    }

    [Fact]
    public void Validate_CopiedCellBudget_ShouldRejectEstimatedRows()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Value");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("invalid");
        var errors = new[]
        {
            new ExcelImportError(ExcelImportErrorCode.InvalidInput, "invalid", "Data", 2, 1, "Value")
        };
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
            MaxCopiedCells = 1
        };

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            NpoiFailureWorkbookPreflight.Validate(workbook, errors,
                new Dictionary<string, ExcelSheetImportRequest>(), options));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    [Fact]
    public void Validate_CopiedCellBudget_ShouldAllowBoundary()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Value");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("invalid");
        var errors = new[]
        {
            new ExcelImportError(ExcelImportErrorCode.InvalidInput, "invalid", "Data", 2, 1, "Value")
        };
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
            MaxCopiedCells = 2
        };

        NpoiFailureWorkbookPreflight.Validate(workbook, errors,
            new Dictionary<string, ExcelSheetImportRequest>(), options);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Validate_CopiedPictureBudget_ShouldRejectBeforeBuilder(bool legacy)
    {
        using var workbook = CreateWorkbookWithPictures(2, legacy);
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
            MaxCopiedPictures = 1
        };

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            NpoiFailureWorkbookPreflight.Validate(workbook, Array.Empty<ExcelImportError>(),
                new Dictionary<string, ExcelSheetImportRequest>(), options));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    [Fact]
    public void Validate_CopiedPictureBudget_ShouldAllowBoundary()
    {
        using var workbook = CreateWorkbookWithPictures(1, legacy: false);
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
            MaxCopiedPictures = 1
        };

        NpoiFailureWorkbookPreflight.Validate(workbook, Array.Empty<ExcelImportError>(),
            new Dictionary<string, ExcelSheetImportRequest>(), options);
    }

    [Fact]
    public void Validate_CopiedPictureBytes_ShouldRejectBeforeBuilder()
    {
        using var workbook = CreateWorkbookWithPictures(1, legacy: false);
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
            MaxCopiedPictureBytes = 1
        };

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            NpoiFailureWorkbookPreflight.Validate(workbook, Array.Empty<ExcelImportError>(),
                new Dictionary<string, ExcelSheetImportRequest>(), options));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    [Fact]
    public void Validate_CopiedPictureBytes_ShouldAllowBoundary()
    {
        using var workbook = CreateWorkbookWithPictures(1, legacy: false);
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
            MaxCopiedPictureBytes = 1_000_000
        };

        NpoiFailureWorkbookPreflight.Validate(workbook, Array.Empty<ExcelImportError>(),
            new Dictionary<string, ExcelSheetImportRequest>(), options);
    }

    [Fact]
    public void Validate_EstimatedTargetObjects_ShouldRejectBeforeBuilder()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Value");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("invalid");
        var errors = new[]
        {
            new ExcelImportError(ExcelImportErrorCode.InvalidInput, "invalid", "Data", 2, 1, "Value")
        };
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
            MaxEstimatedTargetObjects = 4
        };

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            NpoiFailureWorkbookPreflight.Validate(workbook, errors,
                new Dictionary<string, ExcelSheetImportRequest>(), options));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    [Fact]
    public void Validate_EstimatedTargetObjects_ShouldAllowBoundary()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Value");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("invalid");
        var errors = new[]
        {
            new ExcelImportError(ExcelImportErrorCode.InvalidInput, "invalid", "Data", 2, 1, "Value")
        };
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
            MaxEstimatedTargetObjects = 5
        };

        NpoiFailureWorkbookPreflight.Validate(workbook, errors,
            new Dictionary<string, ExcelSheetImportRequest>(), options);
    }

    private static IWorkbook CreateWorkbookWithPictures(int count, bool legacy)
    {
        IWorkbook workbook = legacy ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        var pictureBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        var pictureIndex = workbook.AddPicture(pictureBytes, PictureType.PNG);
        var drawing = sheet.CreateDrawingPatriarch();
        for (var index = 0; index < count; index++)
        {
            var anchor = workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = index;
            anchor.Row1 = index;
            drawing.CreatePicture(anchor, pictureIndex);
        }

        return workbook;
    }
}
