using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.ExcelDataReader.Tests;

/// <summary>
/// ExcelDataReader 分批、资源边界和异常契约测试。
/// </summary>
public sealed class ExcelDataReaderBoundaryTest
{
    /// <summary>
    /// 验证缺少表头时分批导入仅交付结构化错误。
    /// </summary>
    [Fact]
    public void ImportBatches_MissingHeaders_ShouldDeliverOnlyStructuredErrors()
    {
        using var source = CreateSheet("People", new[] { "Unknown" }, includeSpacer: false);
        var batches = new List<ExcelImportBatch<Person>>();

        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source,
            new ExcelBatchImportRequest<Person>("People"), batches.Add);

        var batch = Assert.Single(batches);
        Assert.Empty(batch.Items);
        Assert.Contains(batch.Errors, error => error.Code == ExcelImportErrorCode.InvalidHeader);
        Assert.False(summary.IsSuccess);
        Assert.True(summary.ErrorCount >= 2);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证缺少表头时异步分批导入仅交付结构化错误。
    /// </summary>
    [Fact]
    public async Task ImportBatchesAsync_MissingHeaders_ShouldDeliverOnlyStructuredErrors()
    {
        using var source = CreateSheet("People", new[] { "Unknown" }, includeSpacer: false);
        var batches = new List<ExcelImportBatch<Person>>();

        var summary = await new ExcelDataReaderExcelImporter().ImportBatchesAsync(source,
            new ExcelBatchImportRequest<Person>("People"),
            (batch, _) =>
            {
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.Empty(Assert.Single(batches).Items);
        Assert.False(summary.IsSuccess);
        Assert.Contains(batches[0].Errors, error => error.Code == ExcelImportErrorCode.InvalidHeader);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证完整导入和分批导入遵循空行后的零基数据起始位置。
    /// </summary>
    [Fact]
    public void ImportAndBatch_ShouldHonorZeroBasedDataStartAfterSpacer()
    {
        using var source = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: true);
        var importer = new ExcelDataReaderExcelImporter();
        var workbookRequest = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .Sheet<Person>("People", root => root.People,
                sheet => sheet.DataRowStartIndex(2)));

        var result = importer.Import(source, workbookRequest);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Alice", Assert.Single(result.Workbook.People).Name);
        source.Position = 0;
        var batches = new List<ExcelImportBatch<Person>>();
        var summary = importer.ImportBatches(source, new ExcelBatchImportRequest<Person>("People")
        {
            DataRowStartIndex = 2,
            BatchSize = 1
        }, batches.Add);

        Assert.True(summary.IsSuccess);
        Assert.Equal(3, Assert.Single(batches).FirstRowIndex);
        Assert.Equal("Alice", Assert.Single(batches[0].Items).Name);
    }

    /// <summary>
    /// 验证异步分批导入遵循空行后的零基数据起始位置。
    /// </summary>
    [Fact]
    public async Task ImportBatchesAsync_ShouldHonorZeroBasedDataStartAfterSpacer()
    {
        using var source = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: true);
        var batches = new List<ExcelImportBatch<Person>>();

        var summary = await new ExcelDataReaderExcelImporter().ImportBatchesAsync(source,
            new ExcelBatchImportRequest<Person>("People")
            {
                DataRowStartIndex = 2,
                BatchSize = 1
            },
            (batch, _) =>
            {
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.True(summary.IsSuccess);
        Assert.Equal(3, Assert.Single(batches).FirstRowIndex);
        Assert.Equal("Alice", Assert.Single(batches[0].Items).Name);
    }

    /// <summary>
    /// 验证分批导入在回调前拒绝动态列映射。
    /// </summary>
    [Fact]
    public void ImportBatches_DynamicMappings_ShouldBeRejectedBeforeCallback()
    {
        using var dynamicSource = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: false);
        var dynamicBatches = new List<ExcelImportBatch<Person>>();
        var dynamicException = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ExcelDataReaderExcelImporter().ImportBatches(dynamicSource,
                new ExcelBatchImportRequest<Person>("People")
                {
                    MappingConfiguration = new ExcelMappingConfiguration
                    {
                        DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
                        {
                            new ExcelMappingDynamicColumnConfiguration { Key = "region", Title = "Region" }
                        }
                    }
                }, dynamicBatches.Add));

        Assert.Equal("ExcelDataReader", dynamicException.Provider);
        Assert.Equal(BingOfficesOperation.Import, dynamicException.Operation);
        Assert.Equal(BingOfficesStage.Preflight, dynamicException.Stage);
        Assert.Empty(dynamicBatches);
        Assert.True(dynamicSource.CanRead);

    }

    /// <summary>
    /// 验证单 Sheet 动态列可以跨多个批次写回独立字典并回滚重复值。
    /// </summary>
    [Fact]
    public void ImportBatches_DynamicColumns_ShouldIsolateValuesAndRollbackUnique()
    {
        using var source = CreateDynamicSheet(new[]
        {
            ("Alice", "east", "10"),
            ("Bob", "west", "20"),
            ("Carol", "north", "20")
        });
        var batches = new List<ExcelImportBatch<DynamicBatchPerson>>();
        var request = new ExcelBatchImportRequest<DynamicBatchPerson>("People")
        {
            BatchSize = 1,
            ValidationFailureMode = ExcelValidationFailureMode.Continue,
            MappingConfiguration = new ExcelMappingConfiguration
            {
                DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
                {
                    new() { Key = "region", Title = "Region" },
                    new()
                    {
                        Key = "score", Title = "Score", DataTypeName = "int",
                        ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
                        {
                            new() { Name = "unique", IgnoreEmpty = false }
                        }
                    }
                }
            }
        };

        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source, request, batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(3, summary.RowsRead);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Equal(new[] { "Alice", "Bob" }, batches.SelectMany(batch => batch.Items)
            .Select(item => item.Name));
        Assert.Equal("east", batches[0].Items[0].MappedValues["region"]);
        Assert.Equal(10, batches[0].Items[0].MappedValues["score"]);
        Assert.Equal("west", batches[1].Items[0].MappedValues["region"]);
        Assert.Equal(20, batches[1].Items[0].MappedValues["score"]);
        var duplicate = Assert.Single(batches.SelectMany(batch => batch.Errors));
        Assert.Equal(ExcelImportErrorCode.Validation, duplicate.Code);
        Assert.Equal("score", duplicate.PropertyName);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证单 Sheet 异步批量路径可以写回动态字典。
    /// </summary>
    [Fact]
    public async Task ImportBatchesAsync_DynamicColumns_ShouldConvertAndWriteTarget()
    {
        using var source = CreateDynamicSheet(new[] { ("Alice", "east", "10") });
        var batches = new List<ExcelImportBatch<DynamicBatchPerson>>();
        var request = new ExcelBatchImportRequest<DynamicBatchPerson>("People")
        {
            MappingConfiguration = new ExcelMappingConfiguration
            {
                DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
                {
                    new() { Key = "region", Title = "Region" },
                    new() { Key = "score", Title = "Score", DataTypeName = "int" }
                }
            }
        };

        var summary = await new ExcelDataReaderExcelImporter().ImportBatchesAsync(source, request,
            (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.True(summary.IsSuccess);
        var item = Assert.Single(Assert.Single(batches).Items);
        Assert.Equal("east", item.MappedValues["region"]);
        Assert.Equal(10, item.MappedValues["score"]);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Workbook 异步动态列使用请求目标、命名转换器和定义类型。
    /// </summary>
    [Fact]
    public async Task ImportWorkbookBatchesAsync_DynamicColumns_ShouldConvertAndWriteTarget()
    {
        using var source = CreateDynamicSheet(new[] { ("Alice", "east", "10") });
        var batches = new List<ExcelWorkbookImportBatch>();
        var definition = new[]
        {
            new ExcelDynamicColumnDefinition { Key = "region", Title = "Region", ConverterName = "upper-region" },
            new ExcelDynamicColumnDefinition { Key = "score", Title = "Score", DataType = typeof(int) }
        };
        var request = ExcelImport.Workbook<DynamicBatchWorkbook>(workbook => workbook
            .Sheet<DynamicBatchPerson>("People", root => root.Rows, sheet => sheet
                .DynamicColumns(row => row.Target, definition)
                .BatchSize(1)));

        var summary = await new ExcelDataReaderExcelImporter(
            valueConverters: new IExcelValueConverter[] { new UpperRegionConverter() })
            .ImportWorkbookBatchesAsync(source, request, (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.True(summary.IsSuccess, string.Join(";", batches.SelectMany(batch => batch.Errors)
            .Select(error => error.Message)));
        var item = Assert.Single(Assert.Single(batches).Items.Cast<DynamicBatchPerson>());
        Assert.Equal("EAST", item.Target["region"]);
        Assert.Equal(10, item.Target["score"]);
        Assert.Null(item.MappedValues);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Workbook 动态列开启未知表头失败策略时不交付实体。
    /// </summary>
    [Fact]
    public void ImportWorkbookBatches_DynamicColumns_ShouldRejectUnknownHeader()
    {
        using var source = CreateDynamicSheet(new[] { ("Alice", "east", "10") }, includeUnknown: true);
        var batches = new List<ExcelWorkbookImportBatch>();
        var request = ExcelImport.Workbook<DynamicBatchWorkbook>(workbook => workbook
            .Sheet<DynamicBatchPerson>("People", root => root.Rows, sheet => sheet
                .DynamicColumns(row => row.Target,
                    new[] { new ExcelDynamicColumnDefinition { Key = "region", Title = "Region" } })
                .FailOnUnknownDynamicColumns()));

        var summary = new ExcelDataReaderExcelImporter().ImportWorkbookBatches(source, request, batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Empty(batches.SelectMany(batch => batch.Items));
        var errors = batches.SelectMany(batch => batch.Errors).ToArray();
        Assert.Equal(2, errors.Length);
        var error = Assert.Single(errors.Where(item => item.Header == "Unknown"));
        Assert.Equal(ExcelImportErrorCode.InvalidHeader, error.Code);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证动态列单元格复用 XLSX 原生列表校验并保留动态键坐标。
    /// </summary>
    [Fact]
    public void ImportWorkbookBatches_DynamicColumns_ShouldApplyWorkbookValidation()
    {
        using var source = CreateDynamicNativeValidationSheet();
        var batches = new List<ExcelWorkbookImportBatch>();
        var request = ExcelImport.Workbook<DynamicBatchWorkbook>(workbook => workbook
            .ValidationMode(ExcelImportValidationMode.WorkbookRules)
            .Sheet<DynamicBatchPerson>("People", root => root.Rows, sheet => sheet
                .DynamicColumns(row => row.Target,
                    new[] { new ExcelDynamicColumnDefinition { Key = "region", Title = "Region" } })
                .BatchSize(1)
                .Validate(ExcelValidationFailureMode.Continue)));

        var summary = new ExcelDataReaderExcelImporter().ImportWorkbookBatches(source, request, batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(1, summary.Sheets.Single().Summary.SucceededRows);
        var error = Assert.Single(batches.SelectMany(batch => batch.Errors));
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal("region", error.PropertyName);
        Assert.Equal(3, error.RowIndex);
        Assert.Equal(2, error.ColumnIndex);
        Assert.Equal("bad", error.RawValue);
        Assert.Equal("east", Assert.Single(batches.SelectMany(batch => batch.Items)
            .Cast<DynamicBatchPerson>()).Target["region"]);
    }

    /// <summary>
    /// 验证 XLSX 原生显式列表校验在分批导入中保留合法行并报告非法坐标。
    /// </summary>
    [Fact]
    public void ImportBatches_WorkbookValidation_ShouldRejectInvalidListValue()
    {
        using var source = CreateNativeValidationSheet(useNamedList: false);
        var batches = new List<ExcelImportBatch<Person>>();
        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source,
            new ExcelBatchImportRequest<Person>("People")
            {
                BatchSize = 1,
                ValidationMode = ExcelImportValidationMode.WorkbookRules,
                ValidationFailureMode = ExcelValidationFailureMode.Continue
            }, batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(2, summary.RowsRead);
        Assert.Equal(1, summary.SucceededRows);
        var error = Assert.Single(batches.SelectMany(batch => batch.Errors));
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal("People", error.SheetName);
        Assert.Equal(3, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
        Assert.Equal("Bad", error.RawValue);
        Assert.Equal("Alice", Assert.Single(batches.SelectMany(batch => batch.Items)).Name);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Workbook 分批同步和异步路径均支持名称列表及数值规则。
    /// </summary>
    [Fact]
    public async Task ImportWorkbookBatchesAsync_WorkbookValidation_ShouldUseNamedListAndNumbers()
    {
        using var source = CreateNativeValidationWorkbook();
        var batches = new List<ExcelWorkbookImportBatch>();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .ValidationMode(ExcelImportValidationMode.ConfiguredAndWorkbook)
            .Sheet<Person>("People", root => root.People, sheet => sheet
                .BatchSize(1).Validate(ExcelValidationFailureMode.Continue))
            .Sheet<Person>("Other", root => root.People, sheet => sheet
                .BatchSize(1).Validate(ExcelValidationFailureMode.Continue)));

        var summary = await new ExcelDataReaderExcelImporter().ImportWorkbookBatchesAsync(source, request,
            (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.False(summary.IsSuccess);
        Assert.Equal(1, summary.Sheets.Single(sheet => sheet.SheetName == "People").Summary.SucceededRows);
        Assert.Equal(1, summary.Sheets.Single(sheet => sheet.SheetName == "Other").Summary.SucceededRows);
        var errors = batches.SelectMany(batch => batch.Errors).ToArray();
        Assert.Equal(4, errors.Length);
        Assert.Contains(errors, error => error.SheetName == "People"
            && error.Code == ExcelImportErrorCode.WorkbookValidation && error.RowIndex == 3
            && error.ColumnIndex == 1);
        Assert.Contains(errors, error => error.SheetName == "People"
            && error.Code == ExcelImportErrorCode.WorkbookValidation && error.RowIndex == 3
            && error.ColumnIndex == 2);
        Assert.Contains(errors, error => error.SheetName == "Other"
            && error.Code == ExcelImportErrorCode.WorkbookValidation && error.RowIndex == 3
            && error.ColumnIndex == 1);
        Assert.Contains(errors, error => error.SheetName == "Other"
            && error.Code == ExcelImportErrorCode.WorkbookValidation && error.RowIndex == 3
            && error.ColumnIndex == 2);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLS 输入启用 Workbook 原生校验时在预检阶段明确拒绝。
    /// </summary>
    [Fact]
    public void ImportBatches_WorkbookValidationOnXls_ShouldBeRejectedAtPreflight()
    {
        using var source = CreateLegacyValidationSheet();
        var callbackCalled = false;
        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ExcelDataReaderExcelImporter().ImportBatches(source,
                new ExcelBatchImportRequest<Person>("People")
                {
                    ValidationMode = ExcelImportValidationMode.WorkbookRules
                }, _ => callbackCalled = true));

        Assert.Equal("ExcelDataReader", exception.Provider);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.False(callbackCalled);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 原生日期和文本长度规则在固定列批次中执行。
    /// </summary>
    [Fact]
    public void ImportBatches_WorkbookValidation_ShouldApplyDateAndTextLengthRules()
    {
        using var source = CreateDateAndTextLengthSheet();
        var batches = new List<ExcelImportBatch<ValidationPerson>>();
        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source,
            new ExcelBatchImportRequest<ValidationPerson>("People")
            {
                ValidationMode = ExcelImportValidationMode.WorkbookRules,
                ValidationFailureMode = ExcelValidationFailureMode.Continue
            }, batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(1, summary.SucceededRows);
        var errors = batches.SelectMany(batch => batch.Errors).ToArray();
        Assert.Equal(2, errors.Length);
        Assert.Contains(errors, error => error.RowIndex == 3 && error.ColumnIndex == 1
            && error.Code == ExcelImportErrorCode.WorkbookValidation);
        Assert.Contains(errors, error => error.RowIndex == 3 && error.ColumnIndex == 2
            && error.Code == ExcelImportErrorCode.WorkbookValidation);
        Assert.Equal("ok", Assert.Single(batches.SelectMany(batch => batch.Items)).Description);
    }

    /// <summary>
    /// 验证唯一值在多个批次之间持续生效，并保留合法批次实体。
    /// </summary>
    [Fact]
    public void ImportBatches_UniqueValues_ShouldRejectDuplicatesAcrossBatches()
    {
        using var source = CreateUniqueSheet(new[] { ("Alice", "30"), ("Bob", "31"), ("alice", "32") });
        var batches = new List<ExcelImportBatch<UniquePerson>>();

        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source,
            new ExcelBatchImportRequest<UniquePerson>("People")
            {
                BatchSize = 1,
                ValidationFailureMode = ExcelValidationFailureMode.Continue
            }, batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(3, summary.RowsRead);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Equal(1, summary.FailedRows);
        Assert.Equal(new[] { "Alice", "Bob" }, batches.SelectMany(batch => batch.Items)
            .Select(item => item.Name));
        var duplicate = Assert.Single(batches.SelectMany(batch => batch.Errors));
        Assert.Equal(ExcelImportErrorCode.Validation, duplicate.Code);
        Assert.Equal(4, duplicate.RowIndex);
        Assert.Equal("Name", duplicate.PropertyName);
    }

    /// <summary>
    /// 验证当前行其他字段转换失败时唯一值会回滚，后续行可以复用该值。
    /// </summary>
    [Fact]
    public void ImportBatches_UniqueValues_ShouldRollbackFailedRowReservation()
    {
        using var source = CreateUniqueSheet(new[] { ("Alice", "bad"), ("Alice", "30") });
        var batches = new List<ExcelImportBatch<UniquePerson>>();

        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source,
            new ExcelBatchImportRequest<UniquePerson>("People")
            {
                BatchSize = 1,
                ValidationFailureMode = ExcelValidationFailureMode.Continue
            }, batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(1, summary.SucceededRows);
        Assert.Equal("Alice", Assert.Single(batches.SelectMany(batch => batch.Items)).Name);
        Assert.DoesNotContain(batches.SelectMany(batch => batch.Errors),
            error => error.Code == ExcelImportErrorCode.Validation && error.PropertyName == "Name");
    }

    /// <summary>
    /// 验证唯一值资源上限在批次中返回结构化资源错误。
    /// </summary>
    [Fact]
    public void ImportBatches_UniqueValues_ShouldReportResourceLimit()
    {
        using var source = CreateUniqueSheet(new[] { ("Alice", "30"), ("Bob", "31") });
        var batches = new List<ExcelImportBatch<UniquePerson>>();

        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source,
            new ExcelBatchImportRequest<UniquePerson>("People")
            {
                BatchSize = 1,
                ResourceLimits = new ExcelResourceLimits { MaxTrackedUniqueValues = 1 }
            }, batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(1, summary.SucceededRows);
        var resourceError = Assert.Single(batches.SelectMany(batch => batch.Errors));
        Assert.Equal(ExcelImportErrorCode.ResourceLimit, resourceError.Code);
        Assert.Equal(3, resourceError.RowIndex);
    }

    /// <summary>
    /// 验证异步 Workbook 分批导入按工作表维护独立唯一值状态。
    /// </summary>
    [Fact]
    public async Task ImportWorkbookBatchesAsync_UniqueValues_ShouldIsolateSheetsAndRejectDuplicates()
    {
        using var source = CreateUniqueWorkbook();
        var batches = new List<ExcelWorkbookImportBatch>();
        var request = ExcelImport.Workbook<UniqueWorkbook>(workbook => workbook
            .Sheet<UniquePerson>("People", root => root.Rows, sheet => sheet.BatchSize(1))
            .Sheet<UniquePerson>("Other", root => root.Other, sheet => sheet.BatchSize(1)));

        var summary = await new ExcelDataReaderExcelImporter().ImportWorkbookBatchesAsync(source, request,
            (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.False(summary.IsSuccess);
        Assert.Equal(2, summary.Sheets.Single(sheet => sheet.SheetName == "People").Summary.SucceededRows);
        Assert.Equal(1, summary.Sheets.Single(sheet => sheet.SheetName == "Other").Summary.SucceededRows);
        var peopleErrors = batches.Where(batch => batch.SheetName == "People")
            .SelectMany(batch => batch.Errors).ToArray();
        Assert.Contains(peopleErrors, error => error.Code == ExcelImportErrorCode.Validation
            && error.RowIndex == 4);
        Assert.DoesNotContain(batches.Where(batch => batch.SheetName == "Other")
            .SelectMany(batch => batch.Errors), error => error.Code == ExcelImportErrorCode.Validation);
    }

    /// <summary>
    /// 验证工作表、列和单元格超限时不返回部分实体。
    /// </summary>
    [Fact]
    public void Import_ShouldApplySheetColumnAndCellBudgetsWithoutPartialEntities()
    {
        using var source = CreateWorkbookWithTwoSheets();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxSheets = 1 })
            .Sheet<Person>("People", root => root.People));

        var sheetResult = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.False(sheetResult.IsSuccess);
        Assert.Empty(sheetResult.Workbook.People);
        Assert.Contains(sheetResult.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);

        using var columnSource = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: false);
        var columnResult = new ExcelDataReaderExcelImporter().Import(columnSource,
            ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .ResourceLimits(new ExcelResourceLimits { MaxColumnsPerSheet = 1 })
                .Sheet<Person>("People", root => root.People)));
        Assert.False(columnResult.IsSuccess);
        Assert.Empty(columnResult.Workbook.People);
        Assert.Contains(columnResult.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);

        using var cellSource = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: false);
        var cellResult = new ExcelDataReaderExcelImporter().Import(cellSource,
            ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .ResourceLimits(new ExcelResourceLimits { MaxCells = 2 })
                .Sheet<Person>("People", root => root.People)));
        Assert.False(cellResult.IsSuccess);
        Assert.Empty(cellResult.Workbook.People);
        Assert.Contains(cellResult.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);
    }

    /// <summary>
    /// 验证单元格预算覆盖未选工作表及表头之前的单元格。
    /// </summary>
    [Fact]
    public void Import_ShouldCountWorkbookCellsOnUnselectedSheetsAndBeforeHeader()
    {
        using var workbookSource = CreateWorkbookWithTwoSheets();
        var workbookResult = new ExcelDataReaderExcelImporter().Import(workbookSource,
            ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .ResourceLimits(new ExcelResourceLimits { MaxCells = 4 })
                .Sheet<Person>("People", root => root.People)));

        Assert.False(workbookResult.IsSuccess);
        Assert.Empty(workbookResult.Workbook.People);
        Assert.Contains(workbookResult.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);

        using var preambleSource = CreateSheetWithPreamble();
        var preambleResult = new ExcelDataReaderExcelImporter().Import(preambleSource,
            ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .ResourceLimits(new ExcelResourceLimits { MaxCells = 4 })
                .Sheet<Person>("People", root => root.People,
                    sheet => sheet.HeaderRowIndex(1).DataRowStartIndex(2))));

        Assert.False(preambleResult.IsSuccess);
        Assert.Empty(preambleResult.Workbook.People);
        Assert.Contains(preambleResult.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);

        using var boundarySource = CreateSheetWithPreamble();
        var boundaryResult = new ExcelDataReaderExcelImporter().Import(boundarySource,
            ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .ResourceLimits(new ExcelResourceLimits { MaxCells = 5 })
                .Sheet<Person>("People", root => root.People,
                    sheet => sheet.HeaderRowIndex(1).DataRowStartIndex(2))));

        Assert.True(boundaryResult.IsSuccess, string.Join(";", boundaryResult.Errors.Select(error => error.Message)));
        Assert.Single(boundaryResult.Workbook.People);
    }

    /// <summary>
    /// 验证异步导入的单元格预算覆盖未选工作表。
    /// </summary>
    [Fact]
    public async Task ImportAsync_ShouldCountWorkbookCellsOnUnselectedSheets()
    {
        using var source = CreateWorkbookWithTwoSheets();
        var result = await new ExcelDataReaderExcelImporter().ImportAsync(source,
            ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .ResourceLimits(new ExcelResourceLimits { MaxCells = 4 })
                .Sheet<Person>("People", root => root.People)));

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.People);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);
    }

    /// <summary>
    /// 验证唯一值跟踪超限时返回资源错误且不返回部分实体。
    /// </summary>
    [Fact]
    public void Import_UniqueLimit_ShouldReturnResourceErrorWithoutPartialEntities()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People", new[]
        {
            new UniquePerson { Name = "A", Age = 1 },
            new UniquePerson { Name = "B", Age = 2 }
        }));
        using var source = new MemoryStream();
        new Bing.Offices.Exports.NpoiExcelExporter().Export(export, source);
        source.Position = 0;
        var request = ExcelImport.Workbook<UniqueWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxTrackedUniqueValues = 1 })
            .Sheet<UniquePerson>("People", root => root.Rows));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);
    }

    /// <summary>
    /// 验证损坏工作簿在分批导入时产生结构化异常。
    /// </summary>
    [Fact]
    public void ImportBatches_MalformedWorkbook_ShouldUseStructuredProviderException()
    {
        using var source = new MemoryStream(new byte[] { 1, 2, 3, 4, 5 });
        var exception = Assert.Throws<BingOfficesImportException>(() =>
            new ExcelDataReaderExcelImporter().ImportBatches(source,
                new ExcelBatchImportRequest<Person>("People"), _ => { }));

        Assert.Equal("ExcelDataReader", exception.Provider);
        Assert.Equal(BingOfficesErrorCode.ImportFailed, exception.Code);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal(BingOfficesStage.Read, exception.Stage);
        Assert.NotNull(exception.InnerException);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证分批回调异常原样传播。
    /// </summary>
    [Fact]
    public void ImportBatches_CallbackException_ShouldPropagateOriginalException()
    {
        using var source = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: false);
        var expected = new InvalidOperationException("callback failure");

        var actual = Assert.Throws<InvalidOperationException>(() =>
            new ExcelDataReaderExcelImporter().ImportBatches(source,
                new ExcelBatchImportRequest<Person>("People") { BatchSize = 1 },
                _ => throw expected));

        Assert.Same(expected, actual);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证异步批次取消传播并清理临时文件。
    /// </summary>
    [Fact]
    public async Task ImportBatchesAsync_CallbackCancellation_ShouldPropagateAndCleanStaging()
    {
        using var source = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: false);
        using var cancellation = new CancellationTokenSource();
        var stagingDirectory = Path.Combine(Path.GetTempPath(),
            "bing-offices-excel-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(stagingDirectory);

        try
        {
            var importer = new ExcelDataReaderExcelImporter(() => Path.Combine(
                stagingDirectory, Guid.NewGuid().ToString("N") + ".tmp"));

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                importer.ImportBatchesAsync(source,
                    new ExcelBatchImportRequest<Person>("People") { BatchSize = 1 },
                    (batch, token) =>
                    {
                        cancellation.Cancel();
                        token.ThrowIfCancellationRequested();
                        return Task.CompletedTask;
                    }, cancellation.Token));

            Assert.True(source.CanRead);
            Assert.Empty(Directory.GetFiles(stagingDirectory));
        }
        finally
        {
            if (Directory.Exists(stagingDirectory))
                Directory.Delete(stagingDirectory, recursive: true);
        }
    }

    /// <summary>
    /// Workbook 分批导入预取消时不得调用回调且应保持输入流所有权。
    /// </summary>
    [Fact]
    public void ImportWorkbookBatches_PreCanceled_ShouldNotInvokeCallback()
    {
        using var source = CreateWorkbookWithTwoSheets();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var called = false;
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));

        Assert.Throws<OperationCanceledException>(() =>
            new ExcelDataReaderExcelImporter().ImportWorkbookBatches(source, request, _ => called = true,
                cancellation.Token));

        Assert.False(called);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 缺少工作表时应交付带 Sheet 坐标的结构化错误批次。
    /// </summary>
    [Fact]
    public void ImportWorkbookBatches_MissingSheet_ShouldDeliverStructuredError()
    {
        using var source = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: false);
        var batches = new List<ExcelWorkbookImportBatch>();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook =>
            workbook.Sheet<Person>("Missing", root => root.People));

        var summary = new ExcelDataReaderExcelImporter().ImportWorkbookBatches(source, request, batches.Add);

        Assert.False(summary.IsSuccess);
        var batch = Assert.Single(batches);
        Assert.Equal("Missing", batch.SheetName);
        Assert.Contains(batch.Errors, error => error.Code == ExcelImportErrorCode.InvalidHeader);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证工作表资源超限批次保留实际超限工作表上下文。
    /// </summary>
    [Fact]
    public void ImportWorkbookBatches_MaxSheets_ShouldIdentifyExceededSheet()
    {
        using var source = CreateWorkbookWithTwoSheets();
        var batches = new List<ExcelWorkbookImportBatch>();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxSheets = 1 })
            .Sheet<Person>("People", root => root.People)
            .Sheet<Person>("Other", root => root.People));

        var summary = new ExcelDataReaderExcelImporter().ImportWorkbookBatches(source, request, batches.Add);

        var batch = Assert.Single(batches);
        var sheet = Assert.Single(summary.Sheets);
        Assert.Equal("Other", batch.SheetName);
        Assert.Equal(typeof(Person), batch.ItemType);
        Assert.Equal("Other", sheet.SheetName);
        Assert.Equal("Other", Assert.Single(batch.Errors).SheetName);
        Assert.True(summary.ResourceLimitExceeded);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证异步工作表资源超限批次保留实际超限工作表上下文。
    /// </summary>
    [Fact]
    public async Task ImportWorkbookBatchesAsync_MaxSheets_ShouldIdentifyExceededSheet()
    {
        using var source = CreateWorkbookWithTwoSheets();
        var batches = new List<ExcelWorkbookImportBatch>();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxSheets = 1 })
            .Sheet<Person>("People", root => root.People)
            .Sheet<Person>("Other", root => root.People));

        var summary = await new ExcelDataReaderExcelImporter().ImportWorkbookBatchesAsync(source, request,
            (batch, _) =>
            {
                batches.Add(batch);
                return Task.CompletedTask;
            });

        var batch = Assert.Single(batches);
        Assert.Equal("Other", batch.SheetName);
        Assert.Equal("Other", Assert.Single(summary.Sheets).SheetName);
        Assert.Equal("Other", Assert.Single(batch.Errors).SheetName);
        Assert.True(summary.ResourceLimitExceeded);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证名称和索引选择器指向同一物理工作表时在读取前失败。
    /// </summary>
    [Fact]
    public void ImportWorkbookBatches_DuplicatePhysicalSelector_ShouldFailAsConfiguration()
    {
        using var source = CreateWorkbookWithTwoSheets();
        var batches = new List<ExcelWorkbookImportBatch>();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .Sheet<Person>("People", root => root.People)
            .Sheet<Person>(ExcelSheetSelector.ByIndex(0), root => root.People));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ExcelDataReaderExcelImporter().ImportWorkbookBatches(source, request, batches.Add));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Empty(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证异步名称和索引选择器指向同一物理工作表时在读取前失败。
    /// </summary>
    [Fact]
    public async Task ImportWorkbookBatchesAsync_DuplicatePhysicalSelector_ShouldFailAsConfiguration()
    {
        using var source = CreateWorkbookWithTwoSheets();
        var batches = new List<ExcelWorkbookImportBatch>();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .Sheet<Person>("People", root => root.People)
            .Sheet<Person>(ExcelSheetSelector.ByIndex(0), root => root.People));

        var exception = await Assert.ThrowsAsync<BingOfficesConfigurationException>(() =>
            new ExcelDataReaderExcelImporter().ImportWorkbookBatchesAsync(source, request,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Empty(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// Workbook 异步批次回调取消时应原样传播并清理暂存文件。
    /// </summary>
    [Fact]
    public async Task ImportWorkbookBatchesAsync_CallbackCancellation_ShouldPropagateAndCleanStaging()
    {
        using var source = CreateWorkbookWithTwoSheets();
        using var cancellation = new CancellationTokenSource();
        var stagingDirectory = Path.Combine(Path.GetTempPath(),
            "bing-offices-excel-workbook-batch-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(stagingDirectory);
        try
        {
            var importer = new ExcelDataReaderExcelImporter(() => Path.Combine(
                stagingDirectory, Guid.NewGuid().ToString("N") + ".tmp"));
            var request = ExcelImport.Workbook<ImportWorkbook>(workbook =>
                workbook.Sheet<Person>("People", root => root.People));

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                importer.ImportWorkbookBatchesAsync(source, request, (batch, token) =>
                {
                    cancellation.Cancel();
                    token.ThrowIfCancellationRequested();
                    return Task.CompletedTask;
                }, cancellation.Token));

            Assert.True(source.CanRead);
            Assert.Empty(Directory.GetFiles(stagingDirectory));
        }
        finally
        {
            if (Directory.Exists(stagingDirectory))
                Directory.Delete(stagingDirectory, recursive: true);
        }
    }

    /// <summary>
    /// 创建包含多行唯一性数据的测试工作表。
    /// </summary>
    /// <param name="rows">按顺序写入的姓名和年龄文本。</param>
    /// <returns>包含测试工作表内容且位于起始位置的流。</returns>
    private static MemoryStream CreateUniqueSheet((string Name, string Age)[] rows)
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("People");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Name");
        header.CreateCell(1).SetCellValue("Age");
        for (var index = 0; index < rows.Length; index++)
        {
            var row = sheet.CreateRow(index + 1);
            row.CreateCell(0).SetCellValue(rows[index].Name);
            row.CreateCell(1).SetCellValue(rows[index].Age);
        }
        using var output = new MemoryStream();
        workbook.Write(output, false);
        var bytes = output.ToArray();
        workbook.Dispose();
        return new MemoryStream(bytes, writable: false);
    }

    /// <summary>
    /// 创建包含显式列表校验的 XLSX 工作表。
    /// </summary>
    /// <param name="useNamedList">是否使用工作簿名称列表。</param>
    /// <returns>位于起始位置的 XLSX 流。</returns>
    private static MemoryStream CreateNativeValidationSheet(bool useNamedList)
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("People");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Name");
        header.CreateCell(1).SetCellValue("Age");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("Alice");
        sheet.GetRow(1).CreateCell(1).SetCellValue(30);
        sheet.CreateRow(2).CreateCell(0).SetCellValue("Bad");
        sheet.GetRow(2).CreateCell(1).SetCellValue(31);
        var helper = sheet.GetDataValidationHelper();
        IDataValidationConstraint constraint;
        if (useNamedList)
        {
            var lookup = workbook.CreateSheet("Lookup");
            lookup.CreateRow(0).CreateCell(0).SetCellValue("Alice");
            lookup.CreateRow(1).CreateCell(0).SetCellValue("Bob");
            var name = workbook.CreateName();
            name.NameName = "Allowed";
            name.RefersToFormula = "Lookup!$A$1:$A$2";
            constraint = helper.CreateFormulaListConstraint("Allowed");
        }
        else
        {
            constraint = helper.CreateExplicitListConstraint(new[] { "Alice", "Bob" });
        }
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(1, 2, 0, 0)));
        using var output = new MemoryStream();
        workbook.Write(output, false);
        return new MemoryStream(output.ToArray(), writable: false);
    }

    /// <summary>
    /// 创建包含名称列表和数值校验的多工作表 XLSX 工作簿。
    /// </summary>
    /// <returns>位于起始位置的 XLSX 流。</returns>
    private static MemoryStream CreateNativeValidationWorkbook()
    {
        using var workbook = new XSSFWorkbook();
        var lookup = workbook.CreateSheet("Lookup");
        lookup.CreateRow(0).CreateCell(0).SetCellValue("Alice");
        lookup.CreateRow(1).CreateCell(0).SetCellValue("Bob");
        var name = workbook.CreateName();
        name.NameName = "Allowed";
        name.RefersToFormula = "Lookup!$A$1:$A$2";
        foreach (var sheetName in new[] { "People", "Other" })
        {
            var sheet = workbook.CreateSheet(sheetName);
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("Name");
            header.CreateCell(1).SetCellValue("Age");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("Alice");
            sheet.GetRow(1).CreateCell(1).SetCellValue(30);
            sheet.CreateRow(2).CreateCell(0).SetCellValue("Bad");
            sheet.GetRow(2).CreateCell(1).SetCellValue(150);
            var helper = sheet.GetDataValidationHelper();
            sheet.AddValidationData(helper.CreateValidation(helper.CreateFormulaListConstraint("Allowed"),
                new CellRangeAddressList(1, 2, 0, 0)));
            sheet.AddValidationData(helper.CreateValidation(helper.CreateDecimalConstraint(
                OperatorType.BETWEEN, "0", "100"), new CellRangeAddressList(1, 2, 1, 1)));
        }
        using var output = new MemoryStream();
        workbook.Write(output, false);
        return new MemoryStream(output.ToArray(), writable: false);
    }

    /// <summary>
    /// 创建包含原生列表规则的 XLS 工作簿，用于验证 Provider 边界。
    /// </summary>
    /// <returns>位于起始位置的 XLS 流。</returns>
    private static MemoryStream CreateLegacyValidationSheet()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("People");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Name");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("Alice");
        var helper = sheet.GetDataValidationHelper();
        sheet.AddValidationData(helper.CreateValidation(helper.CreateExplicitListConstraint(new[] { "Alice" }),
            new CellRangeAddressList(1, 1, 0, 0)));
        using var output = new MemoryStream();
        workbook.Write(output, false);
        return new MemoryStream(output.ToArray(), writable: false);
    }

    /// <summary>
    /// 创建包含日期和文本长度校验的 XLSX 工作表。
    /// </summary>
    /// <returns>位于起始位置的 XLSX 流。</returns>
    private static MemoryStream CreateDateAndTextLengthSheet()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("People");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Description");
        header.CreateCell(1).SetCellValue("Date");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("ok");
        sheet.GetRow(1).CreateCell(1).SetCellValue(new DateTime(2026, 8, 25));
        sheet.CreateRow(2).CreateCell(0).SetCellValue("too-long");
        sheet.GetRow(2).CreateCell(1).SetCellValue(new DateTime(2026, 8, 26));
        var dateStyle = workbook.CreateCellStyle();
        dateStyle.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd");
        sheet.GetRow(1).GetCell(1).CellStyle = dateStyle;
        sheet.GetRow(2).GetCell(1).CellStyle = dateStyle;
        var helper = sheet.GetDataValidationHelper();
        sheet.AddValidationData(helper.CreateValidation(helper.CreateTextLengthConstraint(
            OperatorType.BETWEEN, "1", "3"), new CellRangeAddressList(1, 2, 0, 0)));
        sheet.AddValidationData(helper.CreateValidation(helper.CreateDateConstraint(
            OperatorType.LESS_OR_EQUAL, "2026-08-25", null, "yyyy-mm-dd"),
            new CellRangeAddressList(1, 2, 1, 1)));
        using var output = new MemoryStream();
        workbook.Write(output, false);
        return new MemoryStream(output.ToArray(), writable: false);
    }

    /// <summary>
    /// 创建两个工作表且每个工作表独立维护唯一值的测试工作簿。
    /// </summary>
    /// <returns>包含唯一性测试数据且位于起始位置的流。</returns>
    private static MemoryStream CreateUniqueWorkbook()
    {
        var workbook = new XSSFWorkbook();
        var people = workbook.CreateSheet("People");
        var other = workbook.CreateSheet("Other");
        foreach (var sheet in new[] { people, other })
        {
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("Name");
            header.CreateCell(1).SetCellValue("Age");
        }
        var peopleRows = new[] { ("Alice", "30"), ("Bob", "31"), ("alice", "32") };
        for (var index = 0; index < peopleRows.Length; index++)
        {
            var row = people.CreateRow(index + 1);
            row.CreateCell(0).SetCellValue(peopleRows[index].Item1);
            row.CreateCell(1).SetCellValue(peopleRows[index].Item2);
        }
        var otherRow = other.CreateRow(1);
        otherRow.CreateCell(0).SetCellValue("Alice");
        otherRow.CreateCell(1).SetCellValue("40");
        using var output = new MemoryStream();
        workbook.Write(output, false);
        var bytes = output.ToArray();
        workbook.Dispose();
        return new MemoryStream(bytes, writable: false);
    }

    /// <summary>
    /// 创建包含指定表头和一行数据的测试工作簿。
    /// </summary>
    /// <param name="name">测试工作表名称。</param>
    /// <param name="headers">按列排列的表头文本。</param>
    /// <param name="includeSpacer">是否在表头与数据行之间保留空行。</param>
    /// <returns>包含测试工作簿内容且位于起始位置的流。</returns>
    private static MemoryStream CreateSheet(string name, string[] headers, bool includeSpacer)
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet(name);
        var header = sheet.CreateRow(0);
        for (var index = 0; index < headers.Length; index++)
            header.CreateCell(index).SetCellValue(headers[index]);
        var dataRow = sheet.CreateRow(includeSpacer ? 2 : 1);
        for (var index = 0; index < headers.Length; index++)
        {
            if (string.Equals(headers[index], "Age", StringComparison.OrdinalIgnoreCase))
                dataRow.CreateCell(index).SetCellValue(30);
            else if (string.Equals(headers[index], "Name", StringComparison.OrdinalIgnoreCase))
                dataRow.CreateCell(index).SetCellValue("Alice");
            else
                dataRow.CreateCell(index).SetCellValue("value");
        }
        var stream = new MemoryStream();
        workbook.Write(stream, false);
        var bytes = stream.ToArray();
        workbook.Dispose();
        stream.Dispose();
        return new MemoryStream(bytes, writable: false);
    }

    /// <summary>
    /// 创建包含固定列和动态列的测试工作表。
    /// </summary>
    /// <param name="rows">按顺序写入的固定列和动态列值。</param>
    /// <param name="includeUnknown">是否追加未声明的动态表头。</param>
    /// <returns>位于起始位置的 XLSX 流。</returns>
    private static MemoryStream CreateDynamicSheet((string Name, string Region, string Score)[] rows,
        bool includeUnknown = false)
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("People");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Name");
        header.CreateCell(1).SetCellValue("Region");
        header.CreateCell(2).SetCellValue("Score");
        if (includeUnknown)
            header.CreateCell(3).SetCellValue("Unknown");
        for (var index = 0; index < rows.Length; index++)
        {
            var row = sheet.CreateRow(index + 1);
            row.CreateCell(0).SetCellValue(rows[index].Name);
            row.CreateCell(1).SetCellValue(rows[index].Region);
            row.CreateCell(2).SetCellValue(rows[index].Score);
            if (includeUnknown)
                row.CreateCell(3).SetCellValue("ignored");
        }
        using var output = new MemoryStream();
        workbook.Write(output, false);
        return new MemoryStream(output.ToArray(), writable: false);
    }

    /// <summary>
    /// 创建包含动态列表校验的测试工作表。
    /// </summary>
    /// <returns>位于起始位置的 XLSX 流。</returns>
    private static MemoryStream CreateDynamicNativeValidationSheet()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("People");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Name");
        header.CreateCell(1).SetCellValue("Region");
        header.CreateCell(2).SetCellValue("Score");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("Alice");
        sheet.GetRow(1).CreateCell(1).SetCellValue("east");
        sheet.GetRow(1).CreateCell(2).SetCellValue("10");
        sheet.CreateRow(2).CreateCell(0).SetCellValue("Bob");
        sheet.GetRow(2).CreateCell(1).SetCellValue("bad");
        sheet.GetRow(2).CreateCell(2).SetCellValue("20");
        var helper = sheet.GetDataValidationHelper();
        sheet.AddValidationData(helper.CreateValidation(
            helper.CreateExplicitListConstraint(new[] { "east", "west" }),
            new CellRangeAddressList(1, 2, 1, 1)));
        using var output = new MemoryStream();
        workbook.Write(output, false);
        return new MemoryStream(output.ToArray(), writable: false);
    }

    /// <summary>
    /// 创建包含两个工作表的测试工作簿。
    /// </summary>
    /// <returns>包含测试工作簿内容且位于起始位置的流。</returns>
    private static MemoryStream CreateWorkbookWithTwoSheets()
    {
        var first = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: false);
        using var firstWorkbook = WorkbookFactory.Create(first);
        var second = firstWorkbook.CreateSheet("Other");
        var header = second.CreateRow(0);
        header.CreateCell(0).SetCellValue("Name");
        header.CreateCell(1).SetCellValue("Age");
        var row = second.CreateRow(1);
        row.CreateCell(0).SetCellValue("Bob");
        row.CreateCell(1).SetCellValue(31);
        var output = new MemoryStream();
        firstWorkbook.Write(output, false);
        var bytes = output.ToArray();
        return new MemoryStream(bytes, writable: false);
    }

    /// <summary>
    /// 创建表头前含说明行的测试工作簿。
    /// </summary>
    /// <returns>包含测试工作簿内容且位于起始位置的流。</returns>
    private static MemoryStream CreateSheetWithPreamble()
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("People");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Report");
        var header = sheet.CreateRow(1);
        header.CreateCell(0).SetCellValue("Name");
        header.CreateCell(1).SetCellValue("Age");
        var row = sheet.CreateRow(2);
        row.CreateCell(0).SetCellValue("Alice");
        row.CreateCell(1).SetCellValue(30);
        var output = new MemoryStream();
        workbook.Write(output, false);
        var bytes = output.ToArray();
        workbook.Dispose();
        output.Dispose();
        return new MemoryStream(bytes, writable: false);
    }

    /// <summary>
    /// 承载人员导入结果的测试工作簿。
    /// </summary>
    private sealed class ImportWorkbook
    {
        /// <summary>
        /// 获取主要工作表导入的人员集合。
        /// </summary>
        public List<Person> People { get; } = new List<Person>();
    }

    /// <summary>
    /// 用于必填姓名映射的人员测试模型。
    /// </summary>
    private sealed class Person
    {
        /// <summary>
        /// 获取或设置测试人员姓名。
        /// </summary>
        [ExcelRequired]
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置测试人员年龄。
        /// </summary>
        public int Age { get; set; }
    }

    /// <summary>
    /// 用于日期和文本长度原生校验的测试模型。
    /// </summary>
    private sealed class ValidationPerson
    {
        /// <summary>
        /// 获取或设置描述文本。
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// 获取或设置日期。
        /// </summary>
        public DateTime Date { get; set; }
    }

    /// <summary>
    /// 用于姓名唯一性校验的人员测试模型。
    /// </summary>
    private sealed class UniquePerson
    {
        /// <summary>
        /// 获取或设置测试人员姓名。
        /// </summary>
        [ExcelUnique]
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置测试人员年龄。
        /// </summary>
        public int Age { get; set; }
    }

    /// <summary>
    /// 承载姓名唯一性校验结果的测试工作簿。
    /// </summary>
    private sealed class UniqueWorkbook
    {
        /// <summary>
        /// 获取工作表导入的数据行集合。
        /// </summary>
        public List<UniquePerson> Rows { get; } = new List<UniquePerson>();
        /// <summary>
        /// 获取第二个工作表导入的数据行集合。
        /// </summary>
        public List<UniquePerson> Other { get; } = new List<UniquePerson>();
    }

    /// <summary>
    /// 用于分批动态列导入的实体模型。
    /// </summary>
    private sealed class DynamicBatchPerson
    {
        /// <summary>
        /// 获取或设置人员姓名。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置映射计划默认的动态字典。
        /// </summary>
        [DynamicColumn]
        public IDictionary<string, object> MappedValues { get; set; }

        /// <summary>
        /// 获取请求指定的动态目标字典。
        /// </summary>
        [ExcelIgnore]
        public IDictionary<string, object> Target { get; } =
            new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 承载动态列分批导入结果的测试工作簿。
    /// </summary>
    private sealed class DynamicBatchWorkbook
    {
        /// <summary>
        /// 获取工作表导入的数据行集合。
        /// </summary>
        public List<DynamicBatchPerson> Rows { get; } = new List<DynamicBatchPerson>();
    }

    /// <summary>
    /// 将动态区域文本转换为大写的命名转换器。
    /// </summary>
    private sealed class UpperRegionConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "upper-region";

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = Convert.ToString(context.Value, context.Culture)?.ToUpperInvariant();
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = null;
            return false;
        }
    }
}
