using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Exports;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Imports;
using Bing.Offices.Validations;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// Excel 导入导出 P0 回归测试。
/// </summary>
public sealed class ExcelP0RegressionTest
{
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
        var attribute = new RangeAttribute(1, 2);
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
        Assert.Contains(outputSheet.MergedRegions, region => region.FirstRow == 1 && region.LastRow == 1);
        Assert.Single(outputSheet.GetDataValidations());
        Assert.Single(outputSheet.GetAllPictureInfos());
        Assert.Equal("__SourceSheet", outputSheet.GetRow(0).GetCell(2).StringCellValue);
        Assert.Equal("__Errors", outputSheet.GetRow(0).GetCell(5).StringCellValue);
        Assert.NotNull(output.GetSheet("_ImportErrors"));
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

    private sealed class ValidationRow
    {
        public string Code { get; set; }
    }

    private sealed class RowsWorkbook<T> where T : class, new()
    {
        public List<T> Rows { get; } = new();
    }
}
