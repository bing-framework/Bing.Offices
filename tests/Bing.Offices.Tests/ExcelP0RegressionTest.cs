using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Imports;
using Microsoft.Extensions.DependencyInjection;
using Bing.Offices.Validations;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// Excel 导入导出 P0 回归测试。
/// </summary>
public sealed class ExcelP0RegressionTest
{
    /// <summary>
    /// 测试 - 四种 ValidationMode 应只执行声明启用的校验来源，并保持 Workbook 校验先于配置校验。
    /// </summary>
    [Theory]
    [InlineData(ExcelImportValidationMode.Disabled, 0, 0)]
    [InlineData(ExcelImportValidationMode.ConfiguredRules, 1, 0)]
    [InlineData(ExcelImportValidationMode.WorkbookRules, 0, 1)]
    [InlineData(ExcelImportValidationMode.ConfiguredAndWorkbook, 0, 1)]
    public void Import_ValidationMode_ShouldSelectConfiguredAndWorkbookRules(
        ExcelImportValidationMode mode, int configuredErrorCount, int workbookErrorCount)
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Amount");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("11");
            var helper = sheet.GetDataValidationHelper();
            var constraint = helper.CreateExplicitListConstraint(new[] { "10" });
            sheet.AddValidationData(helper.CreateValidation(constraint, new CellRangeAddressList(1, 1, 0, 0)));
        }));
        var request = ExcelImport.Workbook<RowsWorkbook<ValidationModeRow>>(builder => builder
            .ValidationMode(mode)
            .Sheet("Data", root => root.Rows, sheet => sheet.Validate(ValidateMode.Continue)));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.Equal(configuredErrorCount + workbookErrorCount, result.Errors.Count);
        Assert.Equal(configuredErrorCount == 0 && workbookErrorCount == 0, result.IsSuccess);
        Assert.Equal(configuredErrorCount, result.Errors.Count(error => error.Code == ExcelImportErrorCode.MaxValue));
        Assert.Equal(workbookErrorCount,
            result.Errors.Count(error => error.Code == ExcelImportErrorCode.WorkbookValidation));
        if (configuredErrorCount == 0 && workbookErrorCount == 1)
            Assert.Equal(ExcelImportErrorCode.WorkbookValidation, Assert.Single(result.Errors).Code);
    }

    /// <summary>
    /// 测试 - Workbook 校验失败的行即使使用 Continue，也不应继续转换并产生额外错误。
    /// </summary>
    [Theory]
    [InlineData(ExcelImportValidationMode.WorkbookRules)]
    [InlineData(ExcelImportValidationMode.ConfiguredAndWorkbook)]
    public void Import_WorkbookValidationFailure_Continue_ShouldSkipMaterialization(
        ExcelImportValidationMode validationMode)
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(ConversionRow.Count));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("not-a-number");
            var helper = sheet.GetDataValidationHelper();
            var constraint = helper.CreateExplicitListConstraint(new[] { "10" });
            sheet.AddValidationData(helper.CreateValidation(constraint, new CellRangeAddressList(1, 1, 0, 0)));
        }));
        var request = ExcelImport.Workbook<RowsWorkbook<ConversionRow>>(builder => builder
            .ValidationMode(validationMode)
            .Sheet("Data", root => root.Rows, sheet => sheet.Validate(ValidateMode.Continue)));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
    }

    /// <summary>
    /// 测试 - Workbook 小于等于校验只应读取 Formula1，不应因 Formula2 为空拒绝合法边界值。
    /// </summary>
    [Fact]
    public void Import_WorkbookLessOrEqual_ShouldUseFormula1Only()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Amount");
            sheet.CreateRow(1).CreateCell(0).SetCellValue(10);
            sheet.CreateRow(2).CreateCell(0).SetCellValue(11);
            var helper = sheet.GetDataValidationHelper();
            var constraint = helper.CreateintConstraint(7, "10", null);
            sheet.AddValidationData(helper.CreateValidation(constraint,
                new CellRangeAddressList(1, 2, 0, 0)));
        }));
        var request = ExcelImport.Workbook<RowsWorkbook<ValidationModeRow>>(builder => builder
            .ValidationMode(ExcelImportValidationMode.WorkbookRules)
            .Sheet("Data", root => root.Rows));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.Single(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal(3, error.RowIndex);
    }

    /// <summary>
    /// 测试 - Workbook 原生规则应尊重允许空单元格设置，并在禁止空值时报告错误。
    /// </summary>
    [Theory]
    [InlineData(true, 0)]
    [InlineData(false, 1)]
    public void Import_WorkbookValidation_EmptyCellAllowed_ShouldControlEmptyValue(
        bool emptyCellAllowed, int expectedErrorCount)
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Amount");
            sheet.CreateRow(1).CreateCell(0).SetCellType(CellType.Blank);
            sheet.GetRow(1).CreateCell(1).SetCellValue("present");
            var helper = sheet.GetDataValidationHelper();
            var constraint = helper.CreateintConstraint(6, "1", null);
            var validation = helper.CreateValidation(constraint, new CellRangeAddressList(1, 1, 0, 0));
            validation.EmptyCellAllowed = emptyCellAllowed;
            sheet.AddValidationData(validation);
        }));
        var request = ExcelImport.Workbook<RowsWorkbook<ValidationModeRow>>(builder => builder
            .ValidationMode(ExcelImportValidationMode.WorkbookRules)
            .Sheet("Data", root => root.Rows));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.Equal(expectedErrorCount, result.Errors.Count);
        Assert.Equal(expectedErrorCount == 0, result.IsSuccess);
    }

    /// <summary>
    /// 测试 - XSSF 和 HSSF 的数值 Workbook 校验应共同支持八种比较操作符。
    /// </summary>
    [Theory]
    [InlineData(0, "5", "10", "7", true)]
    [InlineData(1, "5", "10", "11", true)]
    [InlineData(2, "7", null, "7", true)]
    [InlineData(3, "7", null, "8", true)]
    [InlineData(4, "7", null, "8", true)]
    [InlineData(5, "7", null, "6", true)]
    [InlineData(6, "7", null, "7", true)]
    [InlineData(7, "7", null, "7", true)]
    [InlineData(0, "5", "10", "11", false)]
    [InlineData(1, "5", "10", "7", false)]
    [InlineData(2, "7", null, "8", false)]
    [InlineData(3, "7", null, "7", false)]
    [InlineData(4, "7", null, "7", false)]
    [InlineData(5, "7", null, "7", false)]
    [InlineData(6, "7", null, "6", false)]
    [InlineData(7, "7", null, "8", false)]
    public void Import_WorkbookNumericValidation_ShouldSupportAllOperators(
        int operation, string first, string second, string value, bool expectedValid)
    {
        // Arrange
        var results = new List<bool>();
        foreach (var workbookType in new[] { "xlsx", "xls" })
        {
            using var source = new MemoryStream(CreateValidationWorkbook(workbookType, operation, first, second, value));
            var request = ExcelImport.Workbook<RowsWorkbook<ValidationTypedRow>>(builder => builder
                .ValidationMode(ExcelImportValidationMode.WorkbookRules)
                .Sheet("Data", root => root.Rows));

            // Act
            var result = new NpoiExcelImporter().Import(source, request);
            if (result.IsSuccess != expectedValid)
            {
                source.Position = 0;
                using var reopened = WorkbookFactory.Create(source);
                var validation = Assert.Single(reopened.GetSheet("Data").GetDataValidations());
                throw new Xunit.Sdk.XunitException($"{workbookType}: operator={validation.ValidationConstraint.Operator}, "
                    + $"formula1=[{validation.ValidationConstraint.Formula1}], formula2=[{validation.ValidationConstraint.Formula2}], "
                    + string.Join("; ", result.Errors.Select(error => error.Message)));
            }
            results.Add(result.IsSuccess);
        }

        // Assert
        Assert.Equal(new[] { expectedValid, expectedValid }, results);
    }

    /// <summary>
    /// 测试 - XSSF 和 HSSF 的 Decimal Workbook 校验应支持八种比较操作符。
    /// </summary>
    [Theory]
    [InlineData(OperatorType.BETWEEN, "1.5", "2.5", "2.0", true)]
    [InlineData(OperatorType.NOT_BETWEEN, "1.5", "2.5", "3.0", true)]
    [InlineData(OperatorType.EQUAL, "2.5", null, "2.5", true)]
    [InlineData(OperatorType.NOT_EQUAL, "2.5", null, "2.0", true)]
    [InlineData(OperatorType.GREATER_THAN, "2.5", null, "3.0", true)]
    [InlineData(OperatorType.LESS_THAN, "2.5", null, "2.0", true)]
    [InlineData(OperatorType.GREATER_OR_EQUAL, "2.5", null, "2.5", true)]
    [InlineData(OperatorType.LESS_OR_EQUAL, "2.5", null, "2.5", true)]
    [InlineData(OperatorType.BETWEEN, "1.5", "2.5", "3.0", false)]
    public void Import_WorkbookDecimalValidation_ShouldSupportAllOperators(
        int operation, string first, string second, string value, bool expectedValid)
    {
        // Arrange
        var results = new List<bool>();
        foreach (var workbookType in new[] { "xlsx", "xls" })
        {
            using var source = new MemoryStream(CreateTypedValidationWorkbook(workbookType, "decimal",
                operation, first, second, value));
            var request = ExcelImport.Workbook<RowsWorkbook<ValidationTypedRow>>(builder => builder
                .ValidationMode(ExcelImportValidationMode.WorkbookRules)
                .Sheet("Data", root => root.Rows));

            // Act
            var result = new NpoiExcelImporter().Import(source, request);

            // Assert
            Assert.Equal(expectedValid, result.IsSuccess);
            results.Add(result.IsSuccess);
        }
        Assert.Equal(new[] { expectedValid, expectedValid }, results);
    }

    /// <summary>
    /// 测试 - 日期、时间和文本长度校验应支持单边操作符及区间边界，并在 XLS/XLSX 中一致。
    /// </summary>
    [Theory]
    [InlineData("date", OperatorType.LESS_OR_EQUAL, "2026-08-25", null, "2026-08-25", true)]
    [InlineData("date", OperatorType.LESS_OR_EQUAL, "2026-08-25", null, "2026-08-26", false)]
    [InlineData("time", OperatorType.GREATER_OR_EQUAL, "09:00:00", null, "09:00:00", true)]
    [InlineData("time", OperatorType.GREATER_OR_EQUAL, "09:00:00", null, "08:59:59", false)]
    [InlineData("text-length", OperatorType.BETWEEN, "3", "5", "abcd", true)]
    [InlineData("text-length", OperatorType.BETWEEN, "3", "5", "abcdef", false)]
    public void Import_WorkbookTypedValidation_ShouldSupportSingleBoundAndRange(
        string validationType, int operation, string first, string second, string value, bool expectedValid)
    {
        // Arrange
        var results = new List<bool>();
        foreach (var workbookType in new[] { "xlsx", "xls" })
        {
            using var source = new MemoryStream(CreateTypedValidationWorkbook(workbookType, validationType,
                operation, first, second, value));
            var request = ExcelImport.Workbook<RowsWorkbook<ValidationTypedRow>>(builder => builder
                .ValidationMode(ExcelImportValidationMode.WorkbookRules)
                .Sheet("Data", root => root.Rows));

            // Act
            var result = new NpoiExcelImporter().Import(source, request);

            // Assert
            if (result.IsSuccess != expectedValid)
            {
                source.Position = 0;
                using var reopened = WorkbookFactory.Create(source);
                var validation = Assert.Single(reopened.GetSheet("Data").GetDataValidations());
                var cell = reopened.GetSheet("Data").GetRow(1).GetCell(0);
                throw new Xunit.Sdk.XunitException($"{workbookType}: cellType={cell.CellType}, "
                    + $"cellValue={cell.ToString()}, formula1=[{validation.ValidationConstraint.Formula1}], "
                    + $"formula2=[{validation.ValidationConstraint.Formula2}], "
                    + $"value1=[{(validation.ValidationConstraint as NPOI.HSSF.UserModel.DVConstraint)?.Value1}], "
                    + $"value2=[{(validation.ValidationConstraint as NPOI.HSSF.UserModel.DVConstraint)?.Value2}], "
                    + string.Join("; ", result.Errors.Select(error => error.Message)));
            }
            results.Add(result.IsSuccess);
        }
        Assert.Equal(new[] { expectedValid, expectedValid }, results);
    }

    /// <summary>
    /// 测试 - 日期、时间和文本长度校验的八种比较操作符应在 XLSX/XLS 中保持一致。
    /// </summary>
    [Theory]
    [InlineData("date", OperatorType.BETWEEN, "2026-08-25", "2026-08-26", "2026-08-25", true)]
    [InlineData("date", OperatorType.BETWEEN, "2026-08-25", "2026-08-26", "2026-08-27", false)]
    [InlineData("date", OperatorType.NOT_BETWEEN, "2026-08-25", "2026-08-26", "2026-08-27", true)]
    [InlineData("date", OperatorType.NOT_BETWEEN, "2026-08-25", "2026-08-26", "2026-08-26", false)]
    [InlineData("date", OperatorType.EQUAL, "2026-08-25", null, "2026-08-25", true)]
    [InlineData("date", OperatorType.EQUAL, "2026-08-25", null, "2026-08-26", false)]
    [InlineData("date", OperatorType.NOT_EQUAL, "2026-08-25", null, "2026-08-26", true)]
    [InlineData("date", OperatorType.NOT_EQUAL, "2026-08-25", null, "2026-08-25", false)]
    [InlineData("date", OperatorType.GREATER_THAN, "2026-08-25", null, "2026-08-26", true)]
    [InlineData("date", OperatorType.GREATER_THAN, "2026-08-25", null, "2026-08-25", false)]
    [InlineData("date", OperatorType.LESS_THAN, "2026-08-25", null, "2026-08-24", true)]
    [InlineData("date", OperatorType.LESS_THAN, "2026-08-25", null, "2026-08-25", false)]
    [InlineData("date", OperatorType.GREATER_OR_EQUAL, "2026-08-25", null, "2026-08-25", true)]
    [InlineData("date", OperatorType.GREATER_OR_EQUAL, "2026-08-25", null, "2026-08-24", false)]
    [InlineData("date", OperatorType.LESS_OR_EQUAL, "2026-08-25", null, "2026-08-25", true)]
    [InlineData("date", OperatorType.LESS_OR_EQUAL, "2026-08-25", null, "2026-08-26", false)]
    [InlineData("time", OperatorType.BETWEEN, "09:00:00", "10:00:00", "09:00:00", true)]
    [InlineData("time", OperatorType.BETWEEN, "09:00:00", "10:00:00", "10:00:01", false)]
    [InlineData("time", OperatorType.NOT_BETWEEN, "09:00:00", "10:00:00", "10:00:01", true)]
    [InlineData("time", OperatorType.NOT_BETWEEN, "09:00:00", "10:00:00", "10:00:00", false)]
    [InlineData("time", OperatorType.EQUAL, "09:00:00", null, "09:00:00", true)]
    [InlineData("time", OperatorType.EQUAL, "09:00:00", null, "09:00:01", false)]
    [InlineData("time", OperatorType.NOT_EQUAL, "09:00:00", null, "09:00:01", true)]
    [InlineData("time", OperatorType.NOT_EQUAL, "09:00:00", null, "09:00:00", false)]
    [InlineData("time", OperatorType.GREATER_THAN, "09:00:00", null, "09:00:01", true)]
    [InlineData("time", OperatorType.GREATER_THAN, "09:00:00", null, "09:00:00", false)]
    [InlineData("time", OperatorType.LESS_THAN, "09:00:00", null, "08:59:59", true)]
    [InlineData("time", OperatorType.LESS_THAN, "09:00:00", null, "09:00:00", false)]
    [InlineData("time", OperatorType.GREATER_OR_EQUAL, "09:00:00", null, "09:00:00", true)]
    [InlineData("time", OperatorType.GREATER_OR_EQUAL, "09:00:00", null, "08:59:59", false)]
    [InlineData("time", OperatorType.LESS_OR_EQUAL, "09:00:00", null, "09:00:00", true)]
    [InlineData("time", OperatorType.LESS_OR_EQUAL, "09:00:00", null, "09:00:01", false)]
    [InlineData("text-length", OperatorType.BETWEEN, "3", "5", "abc", true)]
    [InlineData("text-length", OperatorType.BETWEEN, "3", "5", "abcdef", false)]
    [InlineData("text-length", OperatorType.NOT_BETWEEN, "3", "5", "abcdef", true)]
    [InlineData("text-length", OperatorType.NOT_BETWEEN, "3", "5", "abcde", false)]
    [InlineData("text-length", OperatorType.EQUAL, "3", null, "abc", true)]
    [InlineData("text-length", OperatorType.EQUAL, "3", null, "ab", false)]
    [InlineData("text-length", OperatorType.NOT_EQUAL, "3", null, "ab", true)]
    [InlineData("text-length", OperatorType.NOT_EQUAL, "3", null, "abc", false)]
    [InlineData("text-length", OperatorType.GREATER_THAN, "3", null, "abcd", true)]
    [InlineData("text-length", OperatorType.GREATER_THAN, "3", null, "abc", false)]
    [InlineData("text-length", OperatorType.LESS_THAN, "3", null, "ab", true)]
    [InlineData("text-length", OperatorType.LESS_THAN, "3", null, "abc", false)]
    [InlineData("text-length", OperatorType.GREATER_OR_EQUAL, "3", null, "abc", true)]
    [InlineData("text-length", OperatorType.GREATER_OR_EQUAL, "3", null, "ab", false)]
    [InlineData("text-length", OperatorType.LESS_OR_EQUAL, "3", null, "abc", true)]
    [InlineData("text-length", OperatorType.LESS_OR_EQUAL, "3", null, "abcd", false)]
    public void Import_WorkbookTypedValidation_ShouldSupportAllOperators(
        string validationType, int operation, string first, string second, string value, bool expectedValid)
    {
        var results = new List<bool>();
        foreach (var workbookType in new[] { "xlsx", "xls" })
        {
            using var source = new MemoryStream(CreateTypedValidationWorkbook(workbookType, validationType,
                operation, first, second, value));
            var request = ExcelImport.Workbook<RowsWorkbook<ValidationTypedRow>>(builder => builder
                .ValidationMode(ExcelImportValidationMode.WorkbookRules)
                .Sheet("Data", root => root.Rows));

            var result = new NpoiExcelImporter().Import(source, request);

            Assert.Equal(expectedValid, result.IsSuccess);
            if (operation != OperatorType.BETWEEN && operation != OperatorType.NOT_BETWEEN)
            {
                source.Position = 0;
                using var reopened = WorkbookFactory.Create(source);
                var validation = Assert.Single(reopened.GetSheet("Data").GetDataValidations());
                Assert.True(string.IsNullOrWhiteSpace(validation.ValidationConstraint.Formula2),
                    $"{workbookType} 单边操作符不应写入 Formula2。");
            }
            results.Add(result.IsSuccess);
        }
        Assert.Equal(new[] { expectedValid, expectedValid }, results);
    }

    /// <summary>
    /// 测试 - Continue 收集同一行多个 Workbook 错误，StopOnFirstFailure 只报告首个错误。
    /// </summary>
    [Theory]
    [InlineData(ValidateMode.Continue, 2)]
    [InlineData(ValidateMode.StopOnFirstFailure, 1)]
    public void Import_WorkbookValidation_ShouldRespectContinueAndStop(
        ValidateMode validateMode, int expectedErrorCount)
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(ValidationStopRow.Amount));
            sheet.GetRow(0).CreateCell(1).SetCellValue(nameof(ValidationStopRow.Second));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("bad-1");
            sheet.GetRow(1).CreateCell(1).SetCellValue("bad-2");
            var helper = sheet.GetDataValidationHelper();
            sheet.AddValidationData(helper.CreateValidation(
                helper.CreateExplicitListConstraint(new[] { "good-1" }),
                new CellRangeAddressList(1, 1, 0, 0)));
            sheet.AddValidationData(helper.CreateValidation(
                helper.CreateExplicitListConstraint(new[] { "good-2" }),
                new CellRangeAddressList(1, 1, 1, 1)));
        }));
        var request = ExcelImport.Workbook<RowsWorkbook<ValidationStopRow>>(builder => builder
            .ValidationMode(ExcelImportValidationMode.WorkbookRules)
            .Sheet("Data", root => root.Rows, sheet => sheet.Validate(validateMode)));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.Equal(expectedErrorCount, result.Errors.Count);
        Assert.Empty(result.Workbook.Rows);
    }

    /// <summary>
    /// 测试 - 日期指定精确格式时应校验原始文本，不能被已转换的 DateTime 值绕过。
    /// </summary>
    [Fact]
    public void DateValidation_WithExactFormat_ShouldValidateRawTextBeforeConvertedValue()
    {
        // Arrange
        var rule = new DateTimeExcelValidationRule();
        var attribute = new ExcelDateAttribute("yyyy-MM-dd");
        var converted = new DateTime(2026, 8, 25);
        var invalidText = new ExcelValidationContext("2026/08/25", "Data", 2, 1, "Date",
            converted, typeof(DateTime));
        var validText = new ExcelValidationContext("2026-08-25", "Data", 2, 1, "Date",
            converted, typeof(DateTime));

        // Act
        var invalidResult = rule.Validate(attribute, invalidText);
        var validResult = rule.Validate(attribute, validText);

        // Assert
        Assert.False(invalidResult);
        Assert.True(validResult);
    }

    /// <summary>
    /// 测试 - Attribute 显示值映射为整数时，应先转换为目标枚举再写入属性。
    /// </summary>
    [Fact]
    public void Import_EnumValueMapping_ShouldAssignTargetEnum()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(EnumRow.Status));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("启用");
        }));

        // Act
        var result = new NpoiExcelImporter().Import(source, ExcelImport.Workbook<RowsWorkbook<EnumRow>>(
            builder => builder.Sheet("Data", root => root.Rows)));

        // Assert
        var item = Assert.Single(result.Workbook.Rows);
        Assert.Equal(SampleStatus.Enabled, item.Status);
        Assert.Empty(result.Errors);
    }

    /// <summary>
    /// 测试 - 数值 Formatter 应保持 Numeric Cell，并将格式写入 DataFormat，而不是写成格式化文本。
    /// </summary>
    [Fact]
    public void Export_NumericFormatter_ShouldKeepNumericCell()
    {
        // Arrange
        using var destination = new MemoryStream();
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(NumericRow.Amount),
                    Formatter = "0.00"
                }
            }
        };

        // Act
        new NpoiExcelExporter().Export(ExcelExport.Workbook(builder => builder.AddSheet("Sheet1",
            new[] { new NumericRow { Amount = 12.5m } }, sheet => sheet.Mapping(mapping))), destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var cell = workbook.GetSheetAt(0).GetRow(1).GetCell(0);
        Assert.Equal(CellType.Numeric, cell.CellType);
        Assert.Equal(12.5d, cell.NumericCellValue, 8);
        Assert.NotEqual(0, cell.CellStyle.DataFormat);
    }

    /// <summary>
    /// 测试 - Range 校验应使用请求 Culture 解析原始文本。
    /// </summary>
    [Fact]
    public void RangeValidation_FrenchCulture_ShouldParseDecimalText()
    {
        // Arrange
        var attribute = new ExcelRangeAttribute(1, 2);
        var context = new ExcelValidationContext("1,5", "Data", 2, 1, "Amount",
            null, typeof(decimal), null, CultureInfo.GetCultureInfo("fr-FR"));

        // Act
        var valid = new RangeExcelValidationRule().Validate(attribute, context);

        // Assert
        Assert.True(valid);
    }

    /// <summary>
    /// 测试 - 导入错误应包含 Sheet、行列、稳定列键、表头和原始值，并提供结果统计。
    /// </summary>
    [Fact]
    public void Import_Error_ShouldContainStructuredContext()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Orders");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(ConversionRow.Count));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("not-a-number");
        }));

        // Act
        var result = new NpoiExcelImporter().Import(source,
            ExcelImport.Workbook<RowsWorkbook<ConversionRow>>(builder => builder
                .Sheet("Orders", root => root.Rows)));

        // Assert
        var error = Assert.Single(result.Errors);
        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        Assert.Equal(1, result.Errors.Select(item => (item.SheetName, item.RowIndex)).Distinct().Count());
        Assert.Equal(1, result.Errors.Count);
        Assert.Equal("Orders", error.SheetName);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
        Assert.Equal(nameof(ConversionRow.Count), error.ColumnKey);
        Assert.Equal(nameof(ConversionRow.Count), error.Header);
        Assert.Equal("not-a-number", error.RawValue);
    }

    /// <summary>
    /// 测试 - 图片锚点位于数据单元格时，应绑定到 byte[] 属性而不是被当作空文本。
    /// </summary>
    [Fact]
    public void Import_ImageAnchoredCell_ShouldBindBytes()
    {
        // Arrange
        var imageBytes = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x00
        };
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Photo");
            sheet.CreateRow(1).CreateCell(0).SetCellType(CellType.Blank);
            var pictureIndex = workbook.AddPicture(imageBytes, PictureType.PNG);
            var anchor = workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Row1 = 1;
            anchor.Row2 = 2;
            anchor.Col1 = 0;
            anchor.Col2 = 1;
            sheet.CreateDrawingPatriarch().CreatePicture(anchor, pictureIndex);
        }));
        var request = ExcelImport.Workbook<ImageWorkbook>(builder =>
            builder.Sheet("Data", root => root.Rows));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.True(result.IsSuccess, string.Join("; ", result.Errors.Select(error => error.Message)));
        Assert.Equal(imageBytes, Assert.Single(result.Workbook.Rows).Photo);
    }

    /// <summary>
    /// 测试 - 固定图片列应支持 Fail 和集合 All 两种多图片策略。
    /// </summary>
    [Fact]
    public void Import_FixedImageColumnMultiplicity_ShouldApplyConfiguredPolicy()
    {
        // Arrange
        var bytes = CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Photo");
            sheet.CreateRow(1).CreateCell(0).SetCellType(CellType.Blank);
            for (var index = 0; index < 2; index++)
            {
                var pictureIndex = workbook.AddPicture(new byte[] { 1, 2, (byte)(index + 3) }, PictureType.PNG);
                var anchor = workbook.GetCreationHelper().CreateClientAnchor();
                anchor.Row1 = 1;
                anchor.Row2 = 2;
                anchor.Col1 = 0;
                anchor.Col2 = 1;
                sheet.CreateDrawingPatriarch().CreatePicture(anchor, pictureIndex);
            }
        });
        var failRequest = ExcelImport.Workbook<ImageWorkbook>(builder =>
            builder.Sheet("Data", root => root.Rows, sheet => sheet.Mapping(new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    new() { PropertyName = nameof(ImageRow.Photo), ImageMultiplicity = ExcelImageMultiplicityPolicy.Fail }
                }
            })));
        var allRequest = ExcelImport.Workbook<ImageCollectionWorkbook>(builder =>
            builder.Sheet("Data", root => root.Rows, sheet => sheet.Mapping(new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    new() { PropertyName = nameof(ImageCollectionRow.Photos), Title = "Photo",
                        ImageMultiplicity = ExcelImageMultiplicityPolicy.All }
                }
            })));

        // Act
        using var failSource = new MemoryStream(bytes, writable: false);
        var failResult = new NpoiExcelImporter().Import(failSource, failRequest);
        using var allSource = new MemoryStream(bytes, writable: false);
        var allResult = new NpoiExcelImporter().Import(allSource, allRequest);

        // Assert
        Assert.False(failResult.IsSuccess);
        Assert.Contains(failResult.Errors, error => error.Message.Contains("多个图片", StringComparison.Ordinal));
        Assert.True(allResult.IsSuccess, string.Join("; ", allResult.Errors.Select(error => error.Message)));
        var images = Assert.Single(allResult.Workbook.Rows).Photos;
        Assert.Equal(2, images.Count);
    }

    /// <summary>
    /// 测试 - 不支持的 Workbook 规则在 Report 模式应记录错误但继续导入，Fail 模式应拒绝行。
    /// </summary>
    [Fact]
    public void Import_UnsupportedWorkbookValidation_ShouldReportOrFailByPolicy()
    {
        // Arrange
        var bytes = CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("A");
            var helper = sheet.GetDataValidationHelper();
            var constraint = helper.CreateCustomConstraint("INDIRECT(A1)");
            sheet.AddValidationData(helper.CreateValidation(constraint, new CellRangeAddressList(1, 1, 0, 0)));
        });
        var reportRequest = ExcelImport.Workbook<ValidationWorkbook>(builder =>
            builder.ValidationMode(ExcelImportValidationMode.WorkbookRules)
                .UnsupportedFeaturePolicy(ExcelUnsupportedFeaturePolicy.Report)
                .Sheet("Data", root => root.Rows));
        var failRequest = ExcelImport.Workbook<ValidationWorkbook>(builder =>
            builder.ValidationMode(ExcelImportValidationMode.WorkbookRules)
                .UnsupportedFeaturePolicy(ExcelUnsupportedFeaturePolicy.Fail)
                .Sheet("Data", root => root.Rows));

        // Act
        using var reportSource = new MemoryStream(bytes, writable: false);
        var reportResult = new NpoiExcelImporter().Import(reportSource, reportRequest);
        using var failSource = new MemoryStream(bytes, writable: false);
        var failResult = new NpoiExcelImporter().Import(failSource, failRequest);

        // Assert
        Assert.True(reportResult.Workbook.Rows.Count == 1,
            string.Join("; ", reportResult.Errors.Select(error => error.Message)));
        Assert.Contains(reportResult.Errors, error => error.Code == ExcelImportErrorCode.WorkbookValidation);
        Assert.Empty(failResult.Workbook.Rows);
        Assert.Contains(failResult.Errors, error => error.Code == ExcelImportErrorCode.WorkbookValidation);
    }

    /// <summary>
    /// 测试 - AnnotatedOriginal 失败模式应保留原数据并生成错误批注和汇总 Sheet。
    /// </summary>
    [Fact]
    public void Import_AnnotatedFailureWorkbook_ShouldAnnotateOriginal()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Count");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("invalid");
        }));
        using var failure = new MemoryStream();
        var request = ExcelImport.Workbook<FailureWorkbook>(builder =>
            builder.FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failure
            }).Sheet("Data", root => root.Rows));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.False(result.IsSuccess);
        Assert.Equal(1, result.Errors.Count);
        Assert.True(failure.Length > 0);
        failure.Position = 0;
        using var annotated = WorkbookFactory.Create(failure);
        Assert.NotNull(annotated.GetSheet("_ImportErrors"));
        Assert.NotNull(annotated.GetSheet("Data").GetRow(1).GetCell(0).CellComment);
        Assert.Equal("invalid", annotated.GetSheet("Data").GetRow(1).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 测试 - 失败批注的 Preserve、Append、Replace、Fail 策略应分别遵守冲突契约。
    /// </summary>
    [Theory]
    [InlineData(ExcelImportCommentConflictPolicy.Preserve, "原批注", "source", false)]
    [InlineData(ExcelImportCommentConflictPolicy.Append, "correct format", "source", false)]
    [InlineData(ExcelImportCommentConflictPolicy.Replace, "correct format", "Bing.Offices", false)]
    [InlineData(ExcelImportCommentConflictPolicy.Fail, null, null, true)]
    public void Import_AnnotatedFailureWorkbook_ShouldApplyCommentConflictPolicy(
        ExcelImportCommentConflictPolicy policy, string expectedText, string expectedAuthor, bool shouldThrow)
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Count");
            var cell = sheet.CreateRow(1).CreateCell(0);
            cell.SetCellValue("invalid");
            var anchor = workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = 0;
            anchor.Col2 = 2;
            anchor.Row1 = 1;
            anchor.Row2 = 4;
            var comment = sheet.CreateDrawingPatriarch().CreateCellComment(anchor);
            comment.String = workbook.GetCreationHelper().CreateRichTextString("原批注");
            comment.Author = "source";
            cell.CellComment = comment;
        }));
        using var failure = new MemoryStream();
        var request = ExcelImport.Workbook<FailureWorkbook>(builder => builder
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failure,
                CommentConflictPolicy = policy
            }).Sheet("Data", root => root.Rows));

        // Act
        var action = () => new NpoiExcelImporter().Import(source, request);

        // Assert
        if (shouldThrow)
        {
            Assert.Throws<InvalidOperationException>(action);
            return;
        }
        action();
        failure.Position = 0;
        using var output = WorkbookFactory.Create(failure);
        var text = output.GetSheet("Data").GetRow(1).GetCell(0).CellComment.String.String;
        if (policy == ExcelImportCommentConflictPolicy.Preserve)
            Assert.Equal(expectedText, text);
        else
            Assert.Contains(expectedText, text, StringComparison.Ordinal);
        Assert.Equal(expectedAuthor, output.GetSheet("Data").GetRow(1).GetCell(0).CellComment.Author);
    }

    /// <summary>
    /// 测试 - XLSX 和 XLS ErrorRowsOnly 均应在重开后保留富文本运行格式。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Import_ErrorRowsOnly_ShouldPreserveRichTextRunsAfterReopen(bool legacyFormat)
    {
        // Arrange
        using var source = new MemoryStream(CreateRichTextFailureWorkbook(legacyFormat));
        using var failure = new MemoryStream();
        var request = ExcelImport.Workbook<FailureWorkbook>(builder => builder
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            }).Sheet("Data", root => root.Rows));

        // Act
        new NpoiExcelImporter().Import(source, request);

        // Assert
        failure.Position = 0;
        using var output = WorkbookFactory.Create(failure);
        var richText = output.GetSheet("Data").GetRow(1).GetCell(0).RichStringCellValue;
        Assert.Equal("Error text", richText.String);
        Assert.True(richText.NumFormattingRuns >= 2);
        Assert.True(richText.GetIndexOfFormattingRun(1) > 0);
        Assert.True(GetRichTextFont(output, richText, 0).IsBold);
        Assert.True(GetRichTextFont(output, richText, 1).IsItalic);
    }

    /// <summary>
    /// 测试 - 失败工作簿超过 MaxSerializedBytes 时应在临时输出阶段失败且不污染调用方目标流。
    /// </summary>
    [Fact]
    public void Import_FailureWorkbook_ExceedingMaxSerializedBytes_ShouldNotWriteDestination()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Count");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("invalid");
        }));
        using var failure = new MemoryStream();
        var request = ExcelImport.Workbook<FailureWorkbook>(builder => builder
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failure,
                MaxSerializedBytes = 1
            })
            .Sheet("Data", root => root.Rows));

        // Act
        var action = () => new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.Throws<InvalidOperationException>(action);
        Assert.Equal(0, failure.Length);
    }

    /// <summary>
    /// 测试 - Failure Workbook 应使用请求级临时目录，并在大小失败后清理临时文件。
    /// </summary>
    [Fact]
    public void Import_FailureWorkbook_TemporaryDirectory_ShouldBeCleanedAfterFailure()
    {
        // Arrange
        var temporaryDirectory = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Failure.{Guid.NewGuid():N}");
        Directory.CreateDirectory(temporaryDirectory);
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Count");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("invalid");
        }));
        using var failure = new MemoryStream();
        var request = ExcelImport.Workbook<FailureWorkbook>(builder => builder
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failure,
                MaxSerializedBytes = 1,
                TemporaryDirectory = temporaryDirectory
            })
            .Sheet("Data", root => root.Rows));

        try
        {
            // Act
            Assert.Throws<InvalidOperationException>(() => new NpoiExcelImporter().Import(source, request));

            // Assert
            Assert.Empty(Directory.GetFiles(temporaryDirectory));
        }
        finally
        {
            if (Directory.Exists(temporaryDirectory))
                Directory.Delete(temporaryDirectory, true);
        }
    }

    /// <summary>
    /// 测试 - 清理失败无 sink 时应抛出；sink 抛异常时不应覆盖失败诊断流程。
    /// </summary>
    [Fact]
    public void FailureWorkbook_CleanupFailure_ShouldBeObservableWithoutOverridingSink()
    {
        // Arrange
        var temporaryDirectory = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Failure.{Guid.NewGuid():N}");
        var error = new ExcelImportError(ExcelImportErrorCode.InvalidInput, "invalid", "Data", 2, 1, "Value");
        var errors = new[] { error };
        var resolved = new Dictionary<string, ExcelSheetImportRequest>(StringComparer.OrdinalIgnoreCase);
        var noSinkFileSystem = new FailingDeleteFileSystem();
        using var noSinkDestination = new MemoryStream();
        using var noSinkWorkbook = new XSSFWorkbook();
        noSinkWorkbook.CreateSheet("Data");
        var noSinkOptions = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
            Destination = noSinkDestination,
            TemporaryDirectory = temporaryDirectory
        };

        try
        {
            // Act
            var noSinkException = Assert.Throws<IOException>(() => NpoiFailureWorkbookWriter.Write(noSinkWorkbook,
                noSinkOptions, errors, resolved, CancellationToken.None, noSinkFileSystem));

            // Assert
            Assert.Contains("清理失败", noSinkException.Message);
            Assert.NotNull(noSinkFileSystem.CreatedPath);
            Assert.True(File.Exists(noSinkFileSystem.CreatedPath));

            var sinkCalled = false;
            using var sinkWorkbook = new XSSFWorkbook();
            sinkWorkbook.CreateSheet("Data");
            var sinkOptions = new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = new MemoryStream(),
                TemporaryDirectory = temporaryDirectory,
                DiagnosticSink = diagnostic =>
                {
                    sinkCalled = true;
                    throw new InvalidOperationException("sink failure");
                }
            };
            var sinkFileSystem = new FailingDeleteFileSystem();
            NpoiFailureWorkbookWriter.Write(sinkWorkbook, sinkOptions, errors, resolved, CancellationToken.None,
                sinkFileSystem);
            Assert.True(sinkCalled);
        }
        finally
        {
            if (noSinkFileSystem.CreatedPath != null && File.Exists(noSinkFileSystem.CreatedPath))
                File.Delete(noSinkFileSystem.CreatedPath);
            if (Directory.Exists(temporaryDirectory))
                Directory.Delete(temporaryDirectory, true);
        }
    }

    /// <summary>
    /// 测试 - Failure Workbook 主异常优先时仍应保留清理异常诊断。
    /// </summary>
    [Fact]
    public void FailureWorkbook_PrimaryFailureWithCleanupFailure_ShouldPreserveCleanupException()
    {
        // Arrange
        var fileSystem = new PrimaryAndDeleteFailingFileSystem();
        var error = new ExcelImportError(ExcelImportErrorCode.InvalidInput, "invalid", "Data", 2, 1, "Value");
        using var destination = new MemoryStream();
        using var workbook = new XSSFWorkbook();
        workbook.CreateSheet("Data");
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
            Destination = destination,
            TemporaryDirectory = Path.GetTempPath()
        };

        // Act
        var exception = Assert.Throws<InvalidOperationException>(() => NpoiFailureWorkbookWriter.Write(workbook,
            options, new[] { error }, new Dictionary<string, ExcelSheetImportRequest>(), CancellationToken.None,
            fileSystem));

        // Assert
        Assert.Contains("序列化失败", exception.Message);
        Assert.IsType<IOException>(exception.Data["Bing.Offices.FailureWorkbook.TemporaryCleanupException"]);
    }

    /// <summary>
    /// 测试 - Failure Workbook 目标复制阶段取消时应清理临时文件并传播取消。
    /// </summary>
    [Fact]
    public void FailureWorkbook_CancellationDuringCopy_ShouldCleanTemporaryFile()
    {
        // Arrange
        var temporaryDirectory = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Failure.{Guid.NewGuid():N}");
        Directory.CreateDirectory(temporaryDirectory);
        using var cancellation = new CancellationTokenSource();
        using var destination = new CancelOnWriteStream(cancellation);
        using var workbook = new XSSFWorkbook();
        workbook.CreateSheet("Data");
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
            Destination = destination,
            TemporaryDirectory = temporaryDirectory
        };
        var error = new ExcelImportError(ExcelImportErrorCode.InvalidInput, "invalid", "Data", 2, 1, "Value");

        try
        {
            // Act
            Assert.Throws<OperationCanceledException>(() => NpoiFailureWorkbookWriter.Write(workbook, options,
                new[] { error }, new Dictionary<string, ExcelSheetImportRequest>(), cancellation.Token));

            // Assert
            Assert.Empty(Directory.GetFiles(temporaryDirectory));
        }
        finally
        {
            if (Directory.Exists(temporaryDirectory))
                Directory.Delete(temporaryDirectory, true);
        }
    }

    /// <summary>
    /// 测试 - Failure Workbook 临时目录创建失败时应返回稳定的创建错误并保留原始异常。
    /// </summary>
    [Fact]
    public void FailureWorkbook_TemporaryDirectoryCreationFailure_ShouldClassifyError()
    {
        // Arrange
        using var workbook = new XSSFWorkbook();
        workbook.CreateSheet("Data");
        using var destination = new MemoryStream(new byte[] { 7 });
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
            Destination = destination,
            TemporaryDirectory = "injected"
        };
        var error = new ExcelImportError(ExcelImportErrorCode.InvalidInput, "invalid", "Data", 2, 1, "Value");

        // Act
        var fileSystem = new DirectoryCreationFailingFileSystem();
        var exception = Assert.Throws<IOException>(() => NpoiFailureWorkbookWriter.Write(workbook, options,
            new[] { error }, new Dictionary<string, ExcelSheetImportRequest>(), CancellationToken.None,
            fileSystem));

        // Assert
        Assert.Equal("失败工作簿临时目录创建失败。", exception.Message);
        Assert.IsType<UnauthorizedAccessException>(exception.InnerException);
        Assert.Equal(1, destination.Length);
        Assert.False(fileSystem.DeleteCalled);
    }

    /// <summary>
    /// 测试 - Failure Workbook 临时文件创建失败时应返回稳定的创建错误并保留原始异常。
    /// </summary>
    [Fact]
    public void FailureWorkbook_TemporaryFileCreationFailure_ShouldClassifyError()
    {
        // Arrange
        using var workbook = new XSSFWorkbook();
        workbook.CreateSheet("Data");
        using var destination = new MemoryStream(new byte[] { 7 });
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
            Destination = destination,
            TemporaryDirectory = "injected"
        };
        var error = new ExcelImportError(ExcelImportErrorCode.InvalidInput, "invalid", "Data", 2, 1, "Value");

        // Act
        var fileSystem = new FileCreationFailingFileSystem();
        var exception = Assert.Throws<IOException>(() => NpoiFailureWorkbookWriter.Write(workbook, options,
            new[] { error }, new Dictionary<string, ExcelSheetImportRequest>(), CancellationToken.None,
            fileSystem));

        // Assert
        Assert.Equal("失败工作簿临时文件创建失败。", exception.Message);
        Assert.IsType<IOException>(exception.InnerException);
        Assert.Equal(1, destination.Length);
        Assert.True(fileSystem.DeleteCalled);
        Assert.NotNull(fileSystem.CreatedPath);
    }

    /// <summary>
    /// 测试 - Failure Workbook 复制失败时应保留复制错误分类并清理临时文件。
    /// </summary>
    [Fact]
    public void FailureWorkbook_DestinationCopyFailure_ShouldClassifyErrorAndCleanup()
    {
        // Arrange
        var fileSystem = new TrackingFileSystem();
        using var workbook = new XSSFWorkbook();
        workbook.CreateSheet("Data");
        using var destination = new ThrowingDestinationStream(new byte[] { 7 });
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
            Destination = destination,
            TemporaryDirectory = "injected"
        };
        var error = new ExcelImportError(ExcelImportErrorCode.InvalidInput, "invalid", "Data", 2, 1, "Value");

        // Act
        var exception = Assert.Throws<InvalidOperationException>(() => NpoiFailureWorkbookWriter.Write(workbook, options,
            new[] { error }, new Dictionary<string, ExcelSheetImportRequest>(), CancellationToken.None, fileSystem));

        // Assert
        Assert.Equal("失败工作簿复制到目标流失败。", exception.Message);
        Assert.IsType<IOException>(exception.InnerException);
        Assert.True(fileSystem.DeleteCalled);
        Assert.Equal(1, destination.Length);
        Assert.True(fileSystem.CreatedPath != null);
    }

    /// <summary>
    /// 测试 - ErrorRowsOnly 应从源工作簿独立复制并连续重排失败行及其结构部件。
    /// </summary>
    [Fact]
    public void Import_ErrorRowsOnly_ShouldCopyAndReorderOriginalFailureRows()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("Count");
            sheet.GetRow(2).CreateCell(1).SetCellValue("Formula");
            sheet.CreateRow(3).CreateCell(0).SetCellValue(1);
            var failedRow = sheet.CreateRow(4);
            failedRow.CreateCell(0).SetCellValue("invalid");
            var formula = failedRow.CreateCell(1);
            formula.SetCellFormula("1+1");
            var style = workbook.CreateCellStyle();
            style.DataFormat = workbook.CreateDataFormat().GetFormat("0.00");
            formula.CellStyle = style;
            var commentAnchor = workbook.GetCreationHelper().CreateClientAnchor();
            commentAnchor.Col1 = 0;
            commentAnchor.Col2 = 2;
            commentAnchor.Row1 = 4;
            commentAnchor.Row2 = 7;
            var comment = sheet.CreateDrawingPatriarch().CreateCellComment(commentAnchor);
            comment.String = workbook.GetCreationHelper().CreateRichTextString("源批注");
            comment.Author = "source";
            comment.Visible = true;
            failedRow.GetCell(0).CellComment = comment;
            var hyperlink = workbook.GetCreationHelper().CreateHyperlink(HyperlinkType.Url);
            hyperlink.Address = "https://example.invalid/orders";
            hyperlink.Label = "订单链接";
            failedRow.GetCell(0).Hyperlink = hyperlink;
            sheet.SetColumnWidth(0, 20 * 256);
            sheet.SetColumnHidden(1, true);
            sheet.CreateFreezePane(1, 3);
            sheet.AddMergedRegion(new CellRangeAddress(4, 4, 0, 1));
            var helper = sheet.GetDataValidationHelper();
            var validation = helper.CreateExplicitListConstraint(new[] { "1", "10" });
            sheet.AddValidationData(helper.CreateValidation(validation, new CellRangeAddressList(4, 4, 0, 0)));
            var pictureIndex = workbook.AddPicture(new byte[] { 1, 2, 3 }, PictureType.PNG);
            var anchor = workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Row1 = 4;
            anchor.Row2 = 5;
            anchor.Col1 = 1;
            anchor.Col2 = 2;
            sheet.CreateDrawingPatriarch().CreatePicture(anchor, pictureIndex);
        }));
        using var failure = new MemoryStream();
        var request = ExcelImport.Workbook<FailureWorkbook>(builder =>
            builder.FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            }).Sheet("Data", root => root.Rows, sheet => sheet
                .HeaderRowIndex(2)
                .DataRowStartIndex(3)));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.False(result.IsSuccess);
        failure.Position = 0;
        using var output = WorkbookFactory.Create(failure);
        var outputSheet = output.GetSheet("Data");
        Assert.NotNull(outputSheet);
        Assert.Equal(1, outputSheet.LastRowNum);
        Assert.Equal("invalid", outputSheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal(CellType.Formula, outputSheet.GetRow(1).GetCell(1).CellType);
        Assert.Equal("1+1", outputSheet.GetRow(1).GetCell(1).CellFormula);
        Assert.True(outputSheet.GetRow(1).GetCell(1).CellStyle.DataFormat > 0);
        Assert.Equal("源批注", outputSheet.GetRow(1).GetCell(0).CellComment.String.String);
        Assert.Equal("source", outputSheet.GetRow(1).GetCell(0).CellComment.Author);
        Assert.Equal("https://example.invalid/orders", outputSheet.GetRow(1).GetCell(0).Hyperlink.Address);
        Assert.Equal(20 * 256, outputSheet.GetColumnWidth(0));
        Assert.True(outputSheet.IsColumnHidden(1));
        Assert.True(outputSheet.PaneInformation.IsFreezePane());
        Assert.True(outputSheet.PaneInformation.VerticalSplitPosition > 0);
        Assert.Contains(outputSheet.MergedRegions, region => region.FirstRow == 1 && region.LastRow == 1);
        Assert.Single(outputSheet.GetDataValidations());
        Assert.Single(outputSheet.GetAllPictureInfos());
        Assert.Equal("__SourceSheet", outputSheet.GetRow(0).GetCell(2).StringCellValue);
        Assert.Equal("__Errors", outputSheet.GetRow(0).GetCell(5).StringCellValue);
        Assert.NotNull(output.GetSheet("_ImportErrors"));
    }

    /// <summary>
    /// 测试 - ErrorRowsOnly 按 Sheet 索引解析时应使用实际 Sheet 的非零表头行，而不是源首行。
    /// </summary>
    [Fact]
    public void Import_ErrorRowsOnly_ByIndex_ShouldUseResolvedSheetHeader()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var first = workbook.CreateSheet("Cover");
            first.CreateRow(0).CreateCell(0).SetCellValue("封面");
            var data = workbook.CreateSheet("Data");
            data.CreateRow(0).CreateCell(0).SetCellValue("错误的首行");
            data.CreateRow(2).CreateCell(0).SetCellValue("Count");
            data.CreateRow(3).CreateCell(0).SetCellValue("invalid");
        }));
        using var failure = new MemoryStream();
        var request = ExcelImport.Workbook<FailureWorkbook>(builder => builder
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            })
            .Sheet(ExcelSheetSelector.ByIndex(1), root => root.Rows,
                sheet => sheet.HeaderRowIndex(2).DataRowStartIndex(3)));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.False(result.IsSuccess);
        failure.Position = 0;
        using var output = WorkbookFactory.Create(failure);
        var data = output.GetSheet("Data");
        Assert.Equal("Count", data.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("invalid", data.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("Data", data.GetRow(1).GetCell(1).StringCellValue);
    }

    /// <summary>
    /// 测试 - 输入资源超过请求上限时，应在创建 NPOI Workbook 前拒绝。
    /// </summary>
    [Fact]
    public void Import_InputResourceLimit_ShouldRejectOversizedSource()
    {
        // Arrange
        var bytes = CreateWorkbook(workbook => workbook.CreateSheet("Data"));
        using var source = new MemoryStream(bytes, writable: false);
        var request = ExcelImport.Workbook<FailureWorkbook>(builder =>
            builder.ResourceLimits(new ExcelResourceLimits { MaxInputBytes = 1 })
                .Sheet("Data", root => root.Rows));

        // Act
        var action = () => new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.Throws<InvalidOperationException>(action);
    }

    /// <summary>
    /// 测试 - Workbook 资源限制应在逐行处理和错误收集阶段生效。
    /// </summary>
    [Fact]
    public void Import_WorkbookResourceLimits_ShouldStopRowsAndErrors()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Count");
            sheet.CreateRow(1).CreateCell(0).SetCellValue(1);
            sheet.CreateRow(2).CreateCell(0).SetCellValue("bad");
            sheet.CreateRow(3).CreateCell(0).SetCellValue("bad");
        }));
        var request = ExcelImport.Workbook<FailureWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxRows = 1, MaxErrors = 2 })
            .Sheet("Data", root => root.Rows));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.False(result.IsSuccess);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);
        Assert.Equal(1, result.Errors.Count);
        Assert.False(result.ErrorsTruncated);
        Assert.Equal(2, result.MaxErrors);
        Assert.Equal(1, Assert.Single(result.Workbook.Rows).Count);
    }

    /// <summary>
    /// 测试 - MaxErrors 达到上限后不得追加第 N+1 条错误，并应标记截断元数据。
    /// </summary>
    [Fact]
    public void Import_MaxErrors_ShouldTruncateWithoutExceedingLimit()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Count");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("bad");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("bad");
        }));
        var request = ExcelImport.Workbook<FailureWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxErrors = 1 })
            .Sheet("Data", root => root.Rows));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.False(result.IsSuccess);
        Assert.Single(result.Errors);
        Assert.True(result.ErrorsTruncated);
        Assert.Equal(1, result.MaxErrors);
    }

    /// <summary>
    /// 测试 - 结构错误达到 Workbook 错误上限后，后续 Sheet 应停止处理并标记截断。
    /// </summary>
    [Fact]
    public void Import_StructureErrorAtMaxErrors_ShouldMarkTruncatedAndStopLaterSheets()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var invalid = workbook.CreateSheet("Invalid");
            invalid.CreateRow(0).CreateCell(0).SetCellValue("Unexpected");
            invalid.CreateRow(1).CreateCell(0).SetCellValue("value");
            var valid = workbook.CreateSheet("Valid");
            valid.CreateRow(0).CreateCell(0).SetCellValue("Count");
            valid.CreateRow(1).CreateCell(0).SetCellValue(7);
        }));
        var request = ExcelImport.Workbook<FailureWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxErrors = 1 })
            .Sheet("Invalid", root => root.Rows)
            .Sheet("Valid", root => root.Rows));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.Single(result.Errors);
        Assert.True(result.ErrorsTruncated);
        Assert.Empty(result.Workbook.Rows);
        Assert.Single(result.Sheets);
        Assert.Equal("Invalid", result.Sheets[0].Name);
    }

    /// <summary>
    /// 测试 - 图片数量、单图大小和总大小限制应跨 Sheet 共享，并允许精确达到上限。
    /// </summary>
    [Fact]
    public void Import_WorkbookImageResourceLimits_ShouldBeGlobalAndExact()
    {
        // Arrange
        var bytes = CreateWorkbook(workbook =>
        {
            AddImageSheet(workbook, "First", new byte[] { 1, 2, 3 });
            AddImageSheet(workbook, "Second", new byte[] { 4, 5, 6 });
        });

        // Act
        using var exactSource = new MemoryStream(bytes, writable: false);
        var exactRequest = ExcelImport.Workbook<ImageWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits
            {
                MaxPictures = 2,
                MaxPictureBytes = 3,
                MaxTotalPictureBytes = 6
            })
            .Sheet("First", root => root.Rows)
            .Sheet("Second", root => root.Rows));
        var exactResult = new NpoiExcelImporter().Import(exactSource, exactRequest);

        using var countSource = new MemoryStream(bytes, writable: false);
        var countRequest = ExcelImport.Workbook<ImageWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxPictures = 1 })
            .Sheet("First", root => root.Rows)
            .Sheet("Second", root => root.Rows));
        var countResult = new NpoiExcelImporter().Import(countSource, countRequest);

        using var singleSource = new MemoryStream(bytes, writable: false);
        var singleRequest = ExcelImport.Workbook<ImageWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxPictureBytes = 2 })
            .Sheet("First", root => root.Rows));
        var singleResult = new NpoiExcelImporter().Import(singleSource, singleRequest);

        using var totalSource = new MemoryStream(bytes, writable: false);
        var totalRequest = ExcelImport.Workbook<ImageWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxTotalPictureBytes = 5 })
            .Sheet("First", root => root.Rows)
            .Sheet("Second", root => root.Rows));
        var totalResult = new NpoiExcelImporter().Import(totalSource, totalRequest);

        // Assert
        Assert.True(exactResult.IsSuccess, string.Join("; ", exactResult.Errors.Select(error => error.Message)));
        Assert.Equal(2, exactResult.Workbook.Rows.Count);
        Assert.Single(countResult.Errors);
        Assert.Equal(ExcelImportErrorCode.ResourceLimit, countResult.Errors[0].Code);
        Assert.Single(singleResult.Errors);
        Assert.Equal(ExcelImportErrorCode.ResourceLimit, singleResult.Errors[0].Code);
        Assert.Single(totalResult.Errors);
        Assert.Equal(ExcelImportErrorCode.ResourceLimit, totalResult.Errors[0].Code);
    }

    /// <summary>
    /// 测试 - XLSX/XLS 解析资源边界应在独立进程中可复现，并输出峰值工作集证据。
    /// </summary>
    [Fact]
    public void Import_ResourceProbe_ShouldRunInIndependentProcess()
    {
#if !NET8_0_OR_GREATER
        return;
#else
        // Arrange
        var artifactDirectory = Environment.GetEnvironmentVariable("BING_OFFICES_RESOURCE_PROBE_ARTIFACT");
        var persistArtifacts = !string.IsNullOrWhiteSpace(artifactDirectory);
        var directory = persistArtifacts
            ? Path.GetFullPath(artifactDirectory)
            : Path.Combine(Path.GetTempPath(), $"Bing.Offices.Probe.{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var artifactPath = Path.Combine(directory, "excel-resource-probe-rerun.jsonl");
        if (persistArtifacts)
            File.WriteAllText(artifactPath, string.Empty, Encoding.UTF8);
        var paths = new Dictionary<string, string>
        {
            ["zip"] = Path.Combine(directory, "zip.xlsx"),
            ["dom"] = Path.Combine(directory, "dom.xlsx"),
            ["dom-limit"] = Path.Combine(directory, "dom-limit.xlsx"),
            ["shared-strings"] = Path.Combine(directory, "shared-strings.xlsx"),
            ["styles"] = Path.Combine(directory, "styles.xlsx"),
            ["drawings"] = Path.Combine(directory, "drawings.xlsx"),
            ["ole"] = Path.Combine(directory, "ole.xls")
        };
        File.WriteAllBytes(paths["zip"], CreateProbeWorkbook(4, 2, 0, 0));
        File.WriteAllBytes(paths["dom"], CreateProbeWorkbook(250, 4, 0, 0));
        File.WriteAllBytes(paths["dom-limit"], CreateProbeWorkbook(250, 4, 0, 0));
        File.WriteAllBytes(paths["shared-strings"], CreateProbeWorkbook(20, 20, 400, 0));
        File.WriteAllBytes(paths["styles"], CreateProbeWorkbook(20, 20, 0, 80));
        File.WriteAllBytes(paths["drawings"], CreateProbeWorkbook(10, 4, 0, 0, 6));
        File.WriteAllBytes(paths["ole"], CreateProbeWorkbook<HSSFWorkbook>(4, 2, 0, 0));

        try
        {
            // Act
            var outputs = new Dictionary<string, ProbeOutput>
            {
                ["zip"] = RunResourceProbe(paths["zip"], "zip"),
                ["dom"] = RunResourceProbe(paths["dom"], "dom"),
                ["dom-limit"] = RunResourceProbe(paths["dom-limit"], "dom-limit"),
                ["shared-strings"] = RunResourceProbe(paths["shared-strings"], "shared-strings"),
                ["styles"] = RunResourceProbe(paths["styles"], "styles"),
                ["drawings"] = RunResourceProbe(paths["drawings"], "drawings"),
                ["ole"] = RunResourceProbe(paths["ole"], "ole")
            };
            if (persistArtifacts)
                foreach (var pair in outputs)
                    AppendProbeArtifact(artifactPath, pair.Key, paths[pair.Key], pair.Value);

            // Assert
            Assert.Equal("success", outputs["zip"].Status);
            Assert.Equal("success", outputs["dom"].Status);
            Assert.Equal("resource-limit", outputs["dom-limit"].Status);
            Assert.Equal(250, outputs["dom-limit"].Rows);
            Assert.Equal(100, outputs["dom-limit"].ImportedRows);
            Assert.True(outputs["dom"].Rows > outputs["zip"].Rows);
            Assert.True(outputs["shared-strings"].SharedStrings >= 400);
            Assert.True(outputs["styles"].Styles >= 80);
            Assert.True(outputs["drawings"].Pictures >= 6);
            Assert.Equal("ole", outputs["ole"].Mode);
            Assert.All(outputs.Values, output =>
            {
                Assert.True(output.InputBytes > 0);
                Assert.True(output.ElapsedMilliseconds >= 0);
                Assert.True(output.PeakWorkingSet > 0);
            });
        }
        finally
        {
            if (!persistArtifacts && Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
#endif
    }

    /// <summary>
    /// 测试 - 未绑定图片列时不应扫描图片，也不应触发图片资源限制。
    /// </summary>
    [Fact]
    public void Import_NonImageMapping_ShouldNotScanPictures()
    {
        // Arrange
        var bytes = CreateWorkbook(workbook => AddImageSheet(workbook, "Data", new byte[] { 1, 2, 3 }, "Count"));
        using var source = new MemoryStream(bytes, writable: false);
        var request = ExcelImport.Workbook<FailureWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxPictureBytes = 1 })
            .Sheet("Data", root => root.Rows));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.True(result.IsSuccess, string.Join("; ", result.Errors.Select(error => error.Message)));
        Assert.Single(result.Workbook.Rows);
    }

    /// <summary>
    /// 测试 - Workbook 显式列表校验应在配置模式开启时加入导入结果。
    /// </summary>
    [Fact]
    public void Import_WorkbookExplicitListValidation_ShouldReportInvalidCell()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("A");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("X");
            var helper = sheet.GetDataValidationHelper();
            var constraint = helper.CreateExplicitListConstraint(new[] { "A", "B" });
            var validation = helper.CreateValidation(constraint, new CellRangeAddressList(1, 2, 0, 0));
            sheet.AddValidationData(validation);
        }));
        var request = ExcelImport.Workbook<ValidationWorkbook>(builder =>
            builder.ValidationMode(ExcelImportValidationMode.WorkbookRules)
                .Sheet("Data", root => root.Rows));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.False(result.IsSuccess, string.Join("; ", result.Errors.Select(error => error.Message)));
        Assert.Single(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal(3, error.RowIndex);
    }

    /// <summary>
    /// 测试 - 大范围 Workbook Validation 查询应按矩形区间索引，不为每个单元格创建条目。
    /// </summary>
    [Fact]
    public void ValidationRangeIndex_LargeRectangle_ShouldQueryWithoutCellExpansion()
    {
        // Arrange
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateintConstraint(0, "1", "10");
        var validation = helper.CreateValidation(constraint,
            new CellRangeAddressList(0, 1_000_000, 0, 9));
        var index = ValidationRangeIndex.Create(new[] { validation }, 0, 1_000_000, 0, 9);

        // Act
        var match = index.Get(999_999, 9);
        var outside = index.Get(1_000_001, 0);

        // Assert
        Assert.Single(match);
        Assert.Empty(outside);
    }

    /// <summary>
    /// 测试 - 多个不相邻 Validation 区间查询时应只返回目标行实际覆盖的规则。
    /// </summary>
    [Fact]
    public void ValidationRangeIndex_DisjointRanges_ShouldReturnOnlyMatchingRules()
    {
        // Arrange
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        var helper = sheet.GetDataValidationHelper();
        var validations = new List<IDataValidation>();
        for (var entryNumber = 0; entryNumber < 200; entryNumber++)
        {
            var constraint = helper.CreateintConstraint(0, "1", "10");
            validations.Add(helper.CreateValidation(constraint,
            new CellRangeAddressList(entryNumber * 10, entryNumber * 10, 0, 0)));
        }
        var validationIndex = ValidationRangeIndex.Create(validations, 0, 2_000, 0, 0);

        // Act
        var match = validationIndex.Get(1_000, 0);
        var unrelated = validationIndex.Get(1_001, 0);

        // Assert
        Assert.Single(match);
        Assert.Empty(unrelated);
    }

    /// <summary>
    /// 测试 - 同一行存在大量不相交列范围时，索引查询不应扫描全部规则。
    /// </summary>
    [Fact]
    public void ValidationRangeIndex_OverlappingRows_ShouldLimitColumnCandidates()
    {
        // Arrange
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        var helper = sheet.GetDataValidationHelper();
        var validations = new List<IDataValidation>();
        for (var entryNumber = 0; entryNumber < 200; entryNumber++)
        {
            var constraint = helper.CreateintConstraint(0, "1", "10");
            var column = entryNumber * 2;
            validations.Add(helper.CreateValidation(constraint,
                new CellRangeAddressList(100, 100, column, column)));
        }
        var validationIndex = ValidationRangeIndex.Create(validations, 0, 200, 0, 500);

        // Act
        var match = validationIndex.Get(100, 398, out var candidateChecks);
        var outside = validationIndex.Get(100, 399, out var outsideCandidateChecks);

        // Assert
        Assert.Single(match);
        Assert.Empty(outside);
        Assert.True(candidateChecks < validations.Count);
        Assert.True(outsideCandidateChecks < validations.Count);
    }

    /// <summary>
    /// 测试 - 直接 Cell Range 列表规则应读取引用区域的值，而不是把范围表达式当作常量列表。
    /// </summary>
    [Fact]
    public void Import_DirectCellRangeListValidation_ShouldResolveReferencedValues()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var lookup = workbook.CreateSheet("Lookup");
            lookup.CreateRow(0).CreateCell(0).SetCellValue("A");
            lookup.CreateRow(1).CreateCell(0).SetCellValue("B");
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("A");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("X");
            var helper = sheet.GetDataValidationHelper();
            var constraint = helper.CreateFormulaListConstraint("Lookup!$A$1:$A$2");
            sheet.AddValidationData(helper.CreateValidation(constraint, new CellRangeAddressList(1, 2, 0, 0)));
        }));
        var request = ExcelImport.Workbook<ValidationWorkbook>(builder => builder
            .ValidationMode(ExcelImportValidationMode.WorkbookRules)
            .Sheet("Data", root => root.Rows));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.False(result.IsSuccess);
        Assert.Single(result.Workbook.Rows);
        Assert.Equal(3, Assert.Single(result.Errors).RowIndex);
    }

    private static byte[] CreateWorkbook(Action<IWorkbook> configure)
    {
        using var workbook = new XSSFWorkbook();
        configure(workbook);
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    private static ProbeOutput RunResourceProbe(string inputPath, string mode)
    {
        var probePath = Path.Combine(AppContext.BaseDirectory, "Bing.Offices.ResourceProbe.dll");
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = "dotnet",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            }
        };
        process.StartInfo.ArgumentList.Add(probePath);
        process.StartInfo.ArgumentList.Add(inputPath);
        process.StartInfo.ArgumentList.Add(mode);
        Assert.True(process.Start());
        Assert.True(process.WaitForExit(30000), $"Resource probe timed out: {mode}");
        var output = process.StandardOutput.ReadToEnd();
        var error = process.StandardError.ReadToEnd();
        var exitCode = process.ExitCode;
        Assert.True(exitCode == 0, output + error);
        var values = output.Trim().Split(';')
            .Select(part => part.Split(new[] { '=' }, 2))
            .Where(parts => parts.Length == 2)
            .ToDictionary(parts => parts[0], parts => parts[1], StringComparer.Ordinal);
        return new ProbeOutput(
            values["mode"],
            values["status"],
            long.Parse(values["inputBytes"], CultureInfo.InvariantCulture),
            int.Parse(values["sheets"], CultureInfo.InvariantCulture),
            int.Parse(values["rows"], CultureInfo.InvariantCulture),
            int.Parse(values["columns"], CultureInfo.InvariantCulture),
            int.Parse(values["cells"], CultureInfo.InvariantCulture),
            int.Parse(values["importedRows"], CultureInfo.InvariantCulture),
            ParseUnknownMetric(values["sharedStrings"]),
            ParseUnknownMetric(values["styles"]),
            ParseUnknownMetric(values["pictures"]),
            long.Parse(values["elapsedMs"], CultureInfo.InvariantCulture),
            long.Parse(values["peakWorkingSet"], CultureInfo.InvariantCulture),
            exitCode,
            ComputeSha256(inputPath));
    }

    private static void AppendProbeArtifact(string artifactPath, string scenario, string inputPath,
        ProbeOutput output)
    {
        var line = "{"
            + $"\"scenario\":\"{EscapeJson(scenario)}\","
            + $"\"inputPath\":\"{EscapeJson(inputPath)}\","
            + $"\"inputSha256\":\"{output.InputSha256}\","
            + $"\"inputBytes\":{output.InputBytes},"
            + $"\"mode\":\"{EscapeJson(output.Mode)}\","
            + $"\"status\":\"{EscapeJson(output.Status)}\","
            + $"\"exitCode\":{output.ExitCode},"
            + $"\"sheets\":{output.Sheets},"
            + $"\"rows\":{output.Rows},"
            + $"\"columns\":{output.Columns},"
            + $"\"cells\":{output.Cells},"
            + $"\"importedRows\":{output.ImportedRows},"
            + $"\"sharedStrings\":{output.SharedStrings},"
            + $"\"styles\":{output.Styles},"
            + $"\"pictures\":{output.Pictures},"
            + $"\"elapsedMs\":{output.ElapsedMilliseconds},"
            + $"\"peakWorkingSet\":{output.PeakWorkingSet}"
            + "}";
        File.AppendAllText(artifactPath, line + Environment.NewLine, Encoding.UTF8);
    }

    private static string ComputeSha256(string path)
    {
        using var sha256 = SHA256.Create();
        using var stream = File.OpenRead(path);
        return BitConverter.ToString(sha256.ComputeHash(stream)).Replace("-", string.Empty);
    }

    private static string EscapeJson(string value) => value.Replace("\\", "\\\\")
        .Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n");

    private static int ParseUnknownMetric(string value) => value == "-1"
        ? -1
        : int.Parse(value, CultureInfo.InvariantCulture);

    private static byte[] CreateProbeWorkbook(int rows, int columns, int uniqueStrings, int styles,
        int pictures = 0) => CreateProbeWorkbook<XSSFWorkbook>(rows, columns, uniqueStrings, styles, pictures);

    private static byte[] CreateProbeWorkbook<TWorkbook>(int rows, int columns, int uniqueStrings, int styles,
        int pictures = 0) where TWorkbook : IWorkbook, new()
    {
        return CreateWorkbook<TWorkbook>(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            for (var rowIndex = 0; rowIndex < rows; rowIndex++)
            {
                var row = sheet.CreateRow(rowIndex);
                for (var columnIndex = 0; columnIndex < columns; columnIndex++)
                {
                    var cell = row.CreateCell(columnIndex);
                    cell.SetCellValue(rowIndex == 0
                        ? columnIndex == 0 ? "Name" : $"Extra-{columnIndex}"
                        : uniqueStrings > 0 ? $"shared-{rowIndex}-{columnIndex}" : $"value-{rowIndex}-{columnIndex}");
                    if (styles > 0)
                    {
                        var style = workbook.CreateCellStyle();
                        style.Alignment = (HorizontalAlignment)(style.Index % 3);
                        cell.CellStyle = style;
                    }
                }
            }
            for (var index = 0; index < pictures; index++)
            {
                var anchor = workbook.GetCreationHelper().CreateClientAnchor();
                anchor.Row1 = index;
                anchor.Row2 = index + 1;
                anchor.Col1 = 0;
                anchor.Col2 = 1;
                sheet.CreateDrawingPatriarch().CreatePicture(anchor, workbook.AddPicture(ProbePng, PictureType.PNG));
            }
        });
    }

    private static readonly byte[] ProbePng =
    {
        137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 82,
        0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196, 137,
        0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 248, 207, 192, 240,
        31, 0, 5, 0, 1, 255, 137, 153, 61, 29, 0, 0, 0, 0, 73, 69,
        78, 68, 174, 66, 96, 130
    };

    private sealed class ProbeOutput
    {
        public ProbeOutput(string mode, string status, long inputBytes, int sheets, int rows, int columns,
            int cells, int importedRows, int sharedStrings, int styles, int pictures, long elapsedMilliseconds,
            long peakWorkingSet, int exitCode, string inputSha256)
        {
            Mode = mode;
            Status = status;
            InputBytes = inputBytes;
            Sheets = sheets;
            Rows = rows;
            Columns = columns;
            Cells = cells;
            ImportedRows = importedRows;
            SharedStrings = sharedStrings;
            Styles = styles;
            Pictures = pictures;
            ElapsedMilliseconds = elapsedMilliseconds;
            PeakWorkingSet = peakWorkingSet;
            ExitCode = exitCode;
            InputSha256 = inputSha256;
        }

        public string Mode { get; }
        public string Status { get; }
        public long InputBytes { get; }
        public int Sheets { get; }
        public int Rows { get; }
        public int Columns { get; }
        public int Cells { get; }
        public int ImportedRows { get; }
        public int SharedStrings { get; }
        public int Styles { get; }
        public int Pictures { get; }
        public long ElapsedMilliseconds { get; }
        public long PeakWorkingSet { get; }
        public int ExitCode { get; }
        public string InputSha256 { get; }
    }

    private static byte[] CreateWorkbook<TWorkbook>(Action<TWorkbook> configure)
        where TWorkbook : IWorkbook, new()
    {
        using var workbook = new TWorkbook();
        configure(workbook);
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    private static byte[] CreateValidationWorkbook(string format, int operation, string first,
        string second, string value)
    {
        return format == "xls"
            ? CreateWorkbook<HSSFWorkbook>(workbook => AddNumericValidation(workbook, operation, first, second, value))
            : CreateWorkbook<XSSFWorkbook>(workbook => AddNumericValidation(workbook, operation, first, second, value));
    }

    private static byte[] CreateTypedValidationWorkbook(string format, string validationType, int operation,
        string first, string second, string value)
    {
        return format == "xls"
            ? CreateWorkbook<HSSFWorkbook>(workbook => AddTypedValidation(workbook, validationType, operation,
                first, second, value))
            : CreateWorkbook<XSSFWorkbook>(workbook => AddTypedValidation(workbook, validationType, operation,
                first, second, value));
    }

    private static void AddNumericValidation(IWorkbook workbook, int operation, string first,
        string second, string value)
    {
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Amount");
        sheet.CreateRow(1).CreateCell(0).SetCellValue(value);
        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateintConstraint(operation, first, second);
        sheet.AddValidationData(helper.CreateValidation(constraint, new CellRangeAddressList(1, 1, 0, 0)));
    }

    private static void AddTypedValidation(IWorkbook workbook, string validationType, int operation,
        string first, string second, string value)
    {
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(ValidationTypedRow.Amount));
        sheet.CreateRow(1).CreateCell(0).SetCellValue(value);
        var helper = sheet.GetDataValidationHelper();
        IDataValidationConstraint constraint = validationType switch
        {
            "decimal" => helper.CreateDecimalConstraint(operation, first, second),
            "date" => helper.CreateDateConstraint(operation, first, second, "yyyy-MM-dd"),
            "time" => helper.CreateTimeConstraint(operation, first, second),
            "text-length" => helper.CreateTextLengthConstraint(operation, first, second),
            _ => throw new ArgumentOutOfRangeException(nameof(validationType))
        };
        sheet.AddValidationData(helper.CreateValidation(constraint, new CellRangeAddressList(1, 1, 0, 0)));
    }

    private static byte[] CreateRichTextFailureWorkbook(bool legacyFormat)
    {
        return legacyFormat
            ? CreateWorkbook<HSSFWorkbook>(AddRichTextFailureData)
            : CreateWorkbook<XSSFWorkbook>(AddRichTextFailureData);
    }

    private static void AddRichTextFailureData(IWorkbook workbook)
    {
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Count");
        var cell = sheet.CreateRow(1).CreateCell(0);
        var richText = workbook.GetCreationHelper().CreateRichTextString("Error text");
        for (var index = 0; index < 8; index++)
            workbook.CreateFont().FontName = $"Unused{index}";
        var unusedFont = workbook.CreateFont();
        unusedFont.IsItalic = true;
        var font = workbook.CreateFont();
        font.IsBold = true;
        richText.ApplyFont(0, 5, font);
        richText.ApplyFont(5, 10, unusedFont);
        cell.SetCellValue(richText);
        cell.SetCellValue(richText);
    }

    private static IFont GetRichTextFont(IWorkbook workbook, IRichTextString richText, int run)
    {
        var fontIndex = richText switch
        {
            XSSFRichTextString xssf => xssf.GetFontOfFormattingRun(run).Index,
            HSSFRichTextString hssf => hssf.GetFontOfFormattingRun(run),
            _ => -1
        };
        if (richText is XSSFRichTextString xssfText)
            return xssfText.GetFontOfFormattingRun(run);
        if (fontIndex >= 0)
            return workbook.GetFontAt((short)fontIndex);
        throw new NotSupportedException();
    }

    private static void AddImageSheet(IWorkbook workbook, string name, byte[] pictureBytes,
        string header = "Photo")
    {
        var sheet = workbook.CreateSheet(name);
        sheet.CreateRow(0).CreateCell(0).SetCellValue(header);
        sheet.CreateRow(1).CreateCell(0).SetCellValue(header == "Photo" ? string.Empty : "1");
        var pictureIndex = workbook.AddPicture(pictureBytes, PictureType.PNG);
        var anchor = workbook.GetCreationHelper().CreateClientAnchor();
        anchor.Row1 = 1;
        anchor.Row2 = 2;
        anchor.Col1 = 0;
        anchor.Col2 = 1;
        sheet.CreateDrawingPatriarch().CreatePicture(anchor, pictureIndex);
    }

    private sealed class EnumRow
    {
        [ValueMapping("启用", 1)]
        public SampleStatus Status { get; set; }
    }

    private enum SampleStatus
    {
        Disabled = 0,
        Enabled = 1
    }

    private sealed class NumericRow
    {
        public decimal Amount { get; set; }
    }

    private sealed class ConversionRow
    {
        public int Count { get; set; }
    }

    private sealed class ImageWorkbook
    {
        public List<ImageRow> Rows { get; } = new List<ImageRow>();
    }

    private sealed class ImageRow
    {
        public byte[] Photo { get; set; }
    }

    private sealed class ImageCollectionWorkbook
    {
        public List<ImageCollectionRow> Rows { get; } = new();
    }

    private sealed class ImageCollectionRow
    {
        public IReadOnlyList<ExcelImageData> Photos { get; set; }
    }

    private sealed class FailureWorkbook
    {
        public List<FailureRow> Rows { get; } = new List<FailureRow>();
    }

    private sealed class FailureRow
    {
        public int Count { get; set; }
    }

    private sealed class ValidationWorkbook
    {
        public List<ValidationRow> Rows { get; } = new List<ValidationRow>();
    }

    private sealed class ValidationModeRow
    {
        [ExcelMaxValue(10)]
        public string Amount { get; set; }
    }

    private sealed class ValidationStopRow
    {
        public string Amount { get; set; }
        public string Second { get; set; }
    }

    private sealed class ValidationTypedRow
    {
        public string Amount { get; set; }
    }

    private sealed class ValidationRow
    {
        public string Code { get; set; }
    }

    private sealed class FailingDeleteFileSystem : IFailureWorkbookFileSystem
    {
        public string CreatedPath { get; private set; }

        public void CreateDirectory(string path) => Directory.CreateDirectory(path);

        public Stream CreateFile(string path)
        {
            CreatedPath = path;
            return new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
        }

        public void Delete(string path) => throw new IOException("注入的删除失败");
    }

    private sealed class PrimaryAndDeleteFailingFileSystem : IFailureWorkbookFileSystem
    {
        public void CreateDirectory(string path) { }

        public Stream CreateFile(string path) => new ThrowingWriteStream();

        public void Delete(string path) => throw new IOException("注入的删除失败");
    }

    private sealed class DirectoryCreationFailingFileSystem : IFailureWorkbookFileSystem
    {
        public bool DeleteCalled { get; private set; }

        public void CreateDirectory(string path) => throw new UnauthorizedAccessException("注入的目录创建失败");

        public Stream CreateFile(string path) => throw new InvalidOperationException("不应创建文件");

        public void Delete(string path)
        {
            DeleteCalled = true;
            throw new InvalidOperationException("不应删除文件");
        }
    }

    private sealed class FileCreationFailingFileSystem : IFailureWorkbookFileSystem
    {
        public string CreatedPath { get; private set; }
        public bool DeleteCalled { get; private set; }

        public void CreateDirectory(string path) { }

        public Stream CreateFile(string path)
        {
            CreatedPath = path;
            throw new IOException("注入的文件创建失败");
        }

        public void Delete(string path) => DeleteCalled = true;
    }

    private sealed class TrackingFileSystem : IFailureWorkbookFileSystem
    {
        public string CreatedPath { get; private set; }
        public bool DeleteCalled { get; private set; }

        public void CreateDirectory(string path) { }

        public Stream CreateFile(string path)
        {
            CreatedPath = path;
            return new MemoryStream();
        }

        public void Delete(string path) => DeleteCalled = true;
    }

    private sealed class ThrowingDestinationStream : MemoryStream
    {
        public ThrowingDestinationStream(byte[] buffer) : base(buffer) { }

        public override void Write(byte[] buffer, int offset, int count) => throw new IOException("注入的目标复制失败");
    }

    private sealed class ThrowingWriteStream : MemoryStream
    {
        public override void Write(byte[] buffer, int offset, int count) => throw new IOException("注入的写入失败");
    }

    private sealed class CancelOnWriteStream : MemoryStream
    {
        private readonly CancellationTokenSource _cancellation;

        public CancelOnWriteStream(CancellationTokenSource cancellation)
        {
            _cancellation = cancellation;
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            base.Write(buffer, offset, count);
            _cancellation.Cancel();
        }
    }

    private sealed class RowsWorkbook<T> where T : class, new()
    {
        public List<T> Rows { get; } = new();
    }
}
