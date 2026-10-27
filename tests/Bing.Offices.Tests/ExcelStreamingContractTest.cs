using System;
using System.Collections.Generic;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Xunit;

namespace Bing.Offices.Tests;

public class ExcelStreamingContractTest
{
    [Fact]
    public void Options_ShouldExposeDocumentedDefaults()
    {
        var export = new ExcelStreamingExportOptions();
        Assert.Equal(100, export.RowWindowSize);
        Assert.True(export.CompressTemporaryFiles);
        Assert.Equal(100, export.MaxColumns);
        Assert.Equal(32, export.MaxSheets);
        Assert.Equal(4096, export.MaxStyles);
        Assert.Equal(256, export.MaxFonts);
        Assert.Equal(32767, export.MaxCellCharacters);
        Assert.Equal(262144, export.MaxRowCharacters);
        Assert.Equal(2097152, export.MaxBufferedCharacters);
        Assert.Equal(512L * 1024 * 1024, export.MaxTemporaryBytes);
        Assert.Null(export.TempDirectory);

        var import = new ExcelStreamingImportOptions();
        Assert.Equal(8L * 1024 * 1024, import.SharedStringMemoryBytes);
        Assert.Equal(32767, import.MaxCellCharacters);
        Assert.Equal(262144, import.MaxRowCharacters);
        Assert.Equal(4096, import.MaxStyles);
        Assert.Equal(32, import.MaxSheets);
        Assert.Equal(256, import.MaxFonts);
        Assert.Equal(100000, import.MaxTrackedUniqueValues);
        Assert.Equal(8388608, import.MaxTrackedUniqueCharacters);
        Assert.Equal(1000, import.MaxErrors);
        Assert.Equal(100, import.MaxErrorSamples);
        Assert.Equal(512L * 1024 * 1024, import.MaxTemporaryBytes);
        Assert.Null(import.TempDirectory);

        var resources = import.ResourceLimits;
        Assert.Equal(128L * 1024 * 1024, resources.MaxInputBytes.Value);
        Assert.Null(resources.MaxRows);
        Assert.Null(resources.MaxErrors);
        Assert.Null(resources.MaxPictures);
        Assert.Null(resources.MaxPictureBytes);
        Assert.Null(resources.MaxTotalPictureBytes);
        Assert.Null(resources.MaxTrackedUniqueValues);
        Assert.Equal(StringComparison.OrdinalIgnoreCase, resources.UniqueComparison);
        Assert.Equal(10000, resources.MaxZipEntries.Value);
        Assert.Equal(64L * 1024 * 1024, resources.MaxZipEntryUncompressedBytes.Value);
        Assert.Equal(512L * 1024 * 1024, resources.MaxZipTotalUncompressedBytes.Value);
        Assert.Equal(64L * 1024 * 1024, resources.MaxXmlCharacters.Value);
        Assert.Equal(256, resources.MaxXmlDepth.Value);
        Assert.Equal(1000d, resources.MaxZipCompressionRatio.Value);
        Assert.Equal(64L * 1024 * 1024, resources.MaxSharedStringsBytes.Value);
        Assert.Equal(16L * 1024 * 1024, resources.MaxStylesBytes.Value);
        Assert.Equal(64L * 1024 * 1024, resources.MaxWorksheetBytes.Value);
        Assert.Equal(256L * 1024 * 1024, resources.MaxTotalWorksheetBytes.Value);
    }

    [Fact]
    public void ExportOptions_Validate_ShouldRejectInvalidBudgets()
    {
        Action[] invalid =
        {
            () => new ExcelStreamingExportOptions { RowWindowSize = 0 }.Validate(),
            () => new ExcelStreamingExportOptions { MaxColumns = 0 }.Validate(),
            () => new ExcelStreamingExportOptions { MaxColumns = 16385 }.Validate(),
            () => new ExcelStreamingExportOptions { MaxSheets = 0 }.Validate(),
            () => new ExcelStreamingExportOptions { MaxStyles = 0 }.Validate(),
            () => new ExcelStreamingExportOptions { MaxFonts = 0 }.Validate(),
            () => new ExcelStreamingExportOptions { MaxCellCharacters = 32768 }.Validate(),
            () => new ExcelStreamingExportOptions { MaxRowCharacters = 0 }.Validate(),
            () => new ExcelStreamingExportOptions { MaxBufferedCharacters = 10, MaxRowCharacters = 11 }.Validate(),
            () => new ExcelStreamingExportOptions { MaxTemporaryBytes = 0 }.Validate()
        };

        foreach (var validate in invalid)
            Assert.Throws<ArgumentOutOfRangeException>(validate);
    }

    [Fact]
    public void ImportOptions_Validate_ShouldRejectInvalidBudgetsAndUnboundedPackageLimits()
    {
        Action[] invalid =
        {
            () => new ExcelStreamingImportOptions { SharedStringMemoryBytes = 0 }.Validate(),
            () => new ExcelStreamingImportOptions { MaxCellCharacters = 32768 }.Validate(),
            () => new ExcelStreamingImportOptions { MaxRowCharacters = 0 }.Validate(),
            () => new ExcelStreamingImportOptions { MaxStyles = 0 }.Validate(),
            () => new ExcelStreamingImportOptions { MaxSheets = 0 }.Validate(),
            () => new ExcelStreamingImportOptions { MaxFonts = 0 }.Validate(),
            () => new ExcelStreamingImportOptions { MaxTrackedUniqueValues = 0 }.Validate(),
            () => new ExcelStreamingImportOptions { MaxTrackedUniqueCharacters = 0 }.Validate(),
            () => new ExcelStreamingImportOptions { MaxErrors = 0 }.Validate(),
            () => new ExcelStreamingImportOptions { MaxErrorSamples = -1 }.Validate(),
            () => new ExcelStreamingImportOptions { MaxTemporaryBytes = 0 }.Validate()
        };

        foreach (var validate in invalid)
            Assert.Throws<ArgumentOutOfRangeException>(validate);

        Assert.Throws<ArgumentNullException>(() => new ExcelStreamingImportOptions { ResourceLimits = null }.Validate());
        Assert.Throws<ArgumentException>(() => new ExcelStreamingImportOptions
        {
            ResourceLimits = new ExcelResourceLimits { MaxZipEntries = null }
        }.Validate());
    }

    [Fact]
    public void RowAndSummary_ShouldCopyErrorsAndKeepCompletionSeparateFromSuccess()
    {
        var sourceErrors = new List<ExcelImportError> { CreateError("first") };
        var row = new ExcelImportRow<NpoiExcelStreamingImporterTest.NumberRow>(4,
            new NpoiExcelStreamingImporterTest.NumberRow { Count = 7 }, sourceErrors);

        sourceErrors.Add(CreateError("later"));
        Assert.Equal(4, row.RowIndex);
        Assert.Null(row.Item);
        Assert.False(row.IsValid);
        Assert.Equal("first", Assert.Single(row.Errors).Message);
        Assert.Throws<NotSupportedException>(() => ((IList<ExcelImportError>)row.Errors)[0] = CreateError("mutated"));

        var samples = new List<ExcelImportError> { CreateError("sample") };
        var completedWithErrors = new ExcelStreamingImportSummary(2, 1, 1, 2, 3, samples, ImportStopReason.None);
        samples.Clear();
        Assert.Single(completedWithErrors.Errors);
        Assert.Throws<NotSupportedException>(() => ((IList<ExcelImportError>)completedWithErrors.Errors)[0] = CreateError("mutated"));
        Assert.True(completedWithErrors.ErrorsTruncated);
        Assert.True(completedWithErrors.IsComplete);
        Assert.False(completedWithErrors.IsSuccess);

        var stopped = new ExcelStreamingImportSummary(1, 1, 0, 1, 0, null, ImportStopReason.ConsumerStopped);
        Assert.Empty(stopped.Errors);
        Assert.False(stopped.ErrorsTruncated);
        Assert.False(stopped.IsComplete);
        Assert.False(stopped.IsSuccess);

        var succeeded = new ExcelStreamingImportSummary(1, 1, 0, 1, 0, null, ImportStopReason.None);
        Assert.True(succeeded.IsComplete);
        Assert.True(succeeded.IsSuccess);

        var valid = new ExcelImportRow<NpoiExcelStreamingImporterTest.NumberRow>(0,
            new NpoiExcelStreamingImporterTest.NumberRow { Count = 9 }, null);
        Assert.True(valid.IsValid);
        Assert.Equal(9, valid.Item.Count);
    }

    [Fact]
    public void SheetNameAndIndex_ShouldKeepIndependentConfigurationSnapshots()
    {
        var aliases = new List<string> { "旧标题" };
        var definitions = new List<ExcelDynamicColumnDefinition>
        {
            new() { Key = "extra", Title = "扩展", Aliases = aliases, DataType = typeof(int) }
        };
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(NpoiExcelStreamingImporterTest.DynamicRow.Code), Title = "编码" }
            }
        };

        var byName = ExcelStreamingImport.Sheet<NpoiExcelStreamingImporterTest.DynamicRow>("数据", builder =>
        {
            builder.HeaderRowIndex(2).MaxReadColumns(17).DynamicColumns(row => row.Values, definitions).Mapping(mapping);
        });
        var byIndex = ExcelStreamingImport.Sheet<NpoiExcelStreamingImporterTest.DynamicRow>(3, builder =>
        {
            builder.HeaderRowIndex(2).MaxReadColumns(17).DynamicColumns(row => row.Values, definitions).Mapping(mapping);
        });

        definitions[0] = new ExcelDynamicColumnDefinition { Key = "changed", Title = "已变更" };
        aliases[0] = "新标题";
        mapping.Columns[0].Title = "已变更";
        mapping.Columns.Add(new ExcelColumnConfiguration { PropertyName = "Other", Title = "其他" });

        Assert.Equal(ExcelSheetSelectorKind.ByName, byName.Sheet.Selector.Kind);
        Assert.Equal("数据", byName.Sheet.Selector.Name);
        Assert.Null(byName.Sheet.Selector.Index);
        Assert.Equal(ExcelSheetSelectorKind.ByIndex, byIndex.Sheet.Selector.Kind);
        Assert.Equal(3, byIndex.Sheet.Selector.Index);
        Assert.Null(byIndex.Sheet.Selector.Name);
        Assert.Equal(2, byName.Sheet.HeaderRowIndex);
        Assert.Equal(3, byName.Sheet.DataRowStartIndex);
        Assert.Equal(17, byName.Sheet.MaxReadColumns);

        foreach (var request in new[] { byName, byIndex })
        {
            Assert.Null(request.Sheet.Target);
            var dynamicColumn = Assert.Single(request.Sheet.DynamicColumns);
            Assert.Equal("extra", dynamicColumn.Key);
            Assert.Equal("扩展", dynamicColumn.Title);
            Assert.Equal("旧标题", Assert.Single(dynamicColumn.Aliases));
            Assert.Single(request.Sheet.MappingConfiguration.Columns);
            Assert.Equal("编码", request.Sheet.MappingConfiguration.Columns[0].Title);
        }
    }

    private static ExcelImportError CreateError(string message) =>
        new(ExcelImportErrorCode.ValueConversion, message, "数据", 2, 1, "Count");
}
