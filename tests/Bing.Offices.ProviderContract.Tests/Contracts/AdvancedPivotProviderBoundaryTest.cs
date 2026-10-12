using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Providers;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证未实现高级透视的 Provider 在枚举来源前拒绝请求。
/// </summary>
public sealed class AdvancedPivotProviderBoundaryTest
{
    /// <summary>
    /// 验证不支持的格式与 Provider 在枚举前拒绝高级透视请求。
    /// </summary>
    /// <param name="providerName">待验证的 Provider 名称。</param>
    /// <param name="format">请求的工作簿格式。</param>
    /// <param name="asynchronous">是否调用异步导出入口。</param>
    [Theory]
    [InlineData("NPOI", ExcelFormat.Xls, false)]
    [InlineData("NPOI", ExcelFormat.Xls, true)]
    [InlineData("MiniExcel", ExcelFormat.Xlsx, false)]
    [InlineData("MiniExcel", ExcelFormat.Xlsx, true)]
    [InlineData("SpreadCheetah", ExcelFormat.Xlsx, false)]
    [InlineData("SpreadCheetah", ExcelFormat.Xlsx, true)]
    public async Task AdvancedPivot_ShouldRejectBeforeEnumeratingOrWriting(
        string providerName, ExcelFormat format, bool asynchronous)
    {
        var enumeratedRows = 0;
        IEnumerable<AdvancedPivotProviderRow> Rows()
        {
            enumeratedRows++;
            yield return new AdvancedPivotProviderRow
            {
                Region = "North", Amount = 10, Status = "Open", Channel = "Retail"
            };
        }

        var request = ExcelExport.Workbook(builder => builder
            .Format(format)
            .AddSheet("Data", Rows())
            .PivotTable(new ExcelPivotTableDefinition
            {
                Name = "AdvancedBoundary",
                SourceSheet = "Data",
                SourceRange = new ExcelRangeDefinition
                {
                    StartRow = 0, StartColumn = 0, EndRow = 1, EndColumn = 3
                },
                TargetSheet = "Summary",
                RowFields = new[] { new ExcelPivotAxisFieldDefinition { Field = "Region" } },
                ColumnFields = new[] { new ExcelPivotAxisFieldDefinition { Field = "Status" } },
                ValueFields = new[]
                {
                    new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" },
                    new ExcelPivotValueDefinition
                    {
                        Field = "Amount", Caption = "Count", Aggregation = ExcelPivotAggregation.Count
                    }
                },
                Filters = new[] { new ExcelPivotFilterDefinition { Field = "Channel", Values = new object[] { "Retail" } } }
            }));
        using var destination = new MemoryStream(new byte[] { 17, 23, 42, 91 });
        destination.Position = 2;
        var original = destination.ToArray();
        var position = destination.Position;

        var descriptor = CreateDescriptor(providerName);
        if (providerName == "NPOI")
            Assert.True(descriptor.Supports(ExcelProviderFeatures.AdvancedPivotTables | ExcelProviderFeatures.PivotTables));
        else
        {
            Assert.False(descriptor.Supports(ExcelProviderFeatures.AdvancedPivotTables));
            Assert.False(descriptor.Supports(ExcelProviderFeatures.PivotTables));
        }
        var exception = await Export(descriptor, request, destination, asynchronous);

        Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, exception.Code);
        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(providerName, exception.Provider);
        Assert.Equal(0, enumeratedRows);
        Assert.Equal(original, destination.ToArray());
        Assert.Equal(position, destination.Position);
        Assert.True(destination.CanWrite);
    }

    /// <summary>
    /// 创建指定 Provider 的公开能力描述器。
    /// </summary>
    /// <param name="providerName">Provider 名称。</param>
    /// <returns>用于直接导出和能力检查的 Provider 描述器。</returns>
    private static IExcelProviderFeatureDescriptor CreateDescriptor(string providerName) => providerName switch
    {
        "NPOI" => new NpoiExcelExporter(),
        "MiniExcel" => new MiniExcelExcelExporter(),
        "SpreadCheetah" => new SpreadCheetahStreamingExcelExporter(),
        _ => throw new ArgumentOutOfRangeException(nameof(providerName))
    };

    /// <summary>
    /// 通过 Provider 对应的同步或异步公开入口导出请求。
    /// </summary>
    /// <param name="descriptor">执行导出的 Provider。</param>
    /// <param name="request">包含高级多轴和多度量的请求。</param>
    /// <param name="destination">包含哨兵内容的调用方流。</param>
    /// <param name="asynchronous">是否调用异步入口。</param>
    /// <returns>Provider 抛出的结构化不支持异常。</returns>
    private static async Task<BingOfficesUnsupportedFeatureException> Export(
        IExcelProviderFeatureDescriptor descriptor, ExcelWorkbookExportRequest request,
        Stream destination, bool asynchronous)
    {
        if (descriptor is SpreadCheetahStreamingExcelExporter streamingExporter)
        {
            if (asynchronous)
                return await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() =>
                    streamingExporter.ExportBatchesAsync(request, destination));
            return Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
                streamingExporter.ExportBatches(request, destination));
        }

        var exporter = (IExcelExporter)descriptor;
        if (asynchronous)
            return await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() =>
                exporter.ExportAsync(request, destination));
        return Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            exporter.Export(request, destination));
    }
}
