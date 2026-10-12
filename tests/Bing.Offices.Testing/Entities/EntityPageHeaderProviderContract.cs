using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;

namespace Bing.Offices.Testing.Entities;

/// <summary>验证实体重复页眉的跨 Provider 行为。</summary>
internal static class EntityPageHeaderProviderContract
{
    /// <summary>验证重复页眉的往返与失败边界。</summary>
    internal static void Verify(IExcelEntityExporter exporter, IExcelEntityImporter importer)
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));

        VerifyRoundTrip(exporter, importer, 0);
        VerifyRoundTrip(exporter, importer, 1);
        var multiPage = VerifyRoundTrip(exporter, importer, 5);
        VerifyOccupiedTarget(exporter);
        VerifyMarkerCoordinate(importer, multiPage);
    }

    /// <summary>验证指定明细数量的页眉位置和导入内容。</summary>
    private static byte[] VerifyRoundTrip(IExcelEntityExporter exporter,
        IExcelEntityImporter importer, int count)
    {
        var source = CreateOrder(count);
        var layout = CreateLayout();
        using var output = new MemoryStream();
        exporter.ExportEntity(source, layout, output);
        Require(output.CanWrite && output.Length > 0, "RoundTrip", "导出没有留下完整的调用方流。");
        var bytes = output.ToArray();

        AssertCell(importer, bytes, "A1", "HEAD");
        AssertCell(importer, bytes, "C1", $"1/{Math.Max(1, (count + 1) / 2)}");
        if (count <= 2)
            AssertCell(importer, bytes, "A6", null);
        else
        {
            AssertCell(importer, bytes, "A6", "HEAD");
            AssertCell(importer, bytes, "A11", "HEAD");
            AssertCell(importer, bytes, "C6", "2/3");
            AssertCell(importer, bytes, "C11", "3/3");
            AssertMatchingTitle(importer, bytes, "B3", "B8");
            AssertMatchingTitle(importer, bytes, "C3", "C8");
        }

        using var input = new MemoryStream(bytes, writable: false);
        var result = importer.ImportEntity(input, layout);
        Require(result != null && result.IsSuccess && result.Entity?.Lines != null,
            "RoundTrip", "重复页眉导入失败。");
        Require(result.Entity.Lines.Select(line => line.Code)
                .SequenceEqual(source.Lines.Select(line => line.Code))
            && result.Entity.Lines.Select(line => line.Amount)
                .SequenceEqual(source.Lines.Select(line => line.Amount)),
            "RoundTrip", "页眉被导入为明细或明细内容发生变化。");
        Require(input.CanRead, "RoundTrip", "导入关闭了调用方流。");
        return bytes;
    }

    /// <summary>验证目标页眉不能覆盖已声明的固定单元格。</summary>
    private static void VerifyOccupiedTarget(IExcelEntityExporter exporter)
    {
        var layout = ExcelEntity.Layout<HeaderOrder>(builder => builder
            .Cell("Contract", "A6", order => order.Collision)
            .ListRegion("Contract", "B3", order => order.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD")));
        var source = CreateOrder(3);
        source.Collision = "BUSINESS";
        var original = new byte[] { 4, 8, 2 };
        using var output = new MemoryStream();
        output.Write(original, 0, original.Length);
        output.Position = 0;

        AssertConfigurationFailure(() => exporter.ExportEntity(source, layout, output),
            "OccupiedTarget");
        Require(output.ToArray().SequenceEqual(original) && output.CanWrite,
            "OccupiedTarget", "页眉占用冲突改变了调用方目标流。");
    }

    /// <summary>验证页眉标记不匹配时报告工作表原始坐标。</summary>
    private static void VerifyMarkerCoordinate(IExcelEntityImporter importer, byte[] workbook)
    {
        var layout = CreateLayout("OTHER");
        using var input = new MemoryStream(workbook, writable: false);

        AssertConfigurationFailure(() => importer.ImportEntity(input, layout),
            "MarkerCoordinate", 1, 1);
        Require(input.CanRead, "MarkerCoordinate", "错误导入关闭了调用方来源流。");
    }

    /// <summary>读取指定单元格并核对完整文本。</summary>
    private static void AssertCell(IExcelEntityImporter importer, byte[] workbook,
        string address, string expected)
    {
        Require(ReadCell(importer, workbook, address) == expected,
            "Cell", $"单元格 {address} 内容与预期不一致。");
    }

    /// <summary>核对首页与复制页的列表表头。</summary>
    private static void AssertMatchingTitle(IExcelEntityImporter importer, byte[] workbook,
        string firstAddress, string copiedAddress)
    {
        var title = ReadCell(importer, workbook, firstAddress);
        Require(!string.IsNullOrEmpty(title)
            && title == ReadCell(importer, workbook, copiedAddress),
            "CopiedTitle", $"复制页列表表头 {copiedAddress} 与首页不一致。");
    }

    /// <summary>读取指定单元格的完整文本。</summary>
    private static string ReadCell(IExcelEntityImporter importer, byte[] workbook, string address)
    {
        var layout = ExcelEntity.Layout<CellProbe>(builder => builder
            .Cell("Contract", address, probe => probe.Value));
        using var input = new MemoryStream(workbook, writable: false);
        var result = importer.ImportEntity(input, layout);
        Require(result != null && result.IsSuccess && result.Entity != null,
            "ReadCell", $"无法读取单元格 {address}。");
        return result.Entity.Value;
    }

    /// <summary>检查配置错误的阶段和原始坐标。</summary>
    private static void AssertConfigurationFailure(Action operation, string scenario,
        int? row = null, int? column = null)
    {
        try
        {
            operation();
        }
        catch (BingOfficesConfigurationException exception)
        {
            Require(exception.Code == BingOfficesErrorCode.ConfigurationInvalid
                && exception.Stage == BingOfficesStage.Plan
                && (!row.HasValue || exception.SheetName == "Contract"
                    && exception.RowIndex == row && exception.ColumnIndex == column),
                scenario, $"配置错误不符: Stage={exception.Stage}, Sheet={exception.SheetName}, "
                    + $"Row={exception.RowIndex}, Column={exception.ColumnIndex}; 预期 {row}/{column}。");
            return;
        }
        throw new InvalidOperationException($"Entity PageHeader contract {scenario}: 未拒绝无效布局。");
    }

    /// <summary>创建重复页眉合同布局。</summary>
    private static ExcelEntityLayout<HeaderOrder> CreateLayout(string marker = "HEAD") =>
        ExcelEntity.Layout<HeaderOrder>(builder => builder
            .ListRegion("Contract", "B3", order => order.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", marker, header => header
                    .Cell("C1", page => $"{page.PageNumber}/{page.PageCount}"))));

    /// <summary>创建指定行数的合同明细。</summary>
    private static HeaderOrder CreateOrder(int count) => new()
    {
        Lines = Enumerable.Range(1, count).Select(index => new HeaderLine
        {
            Code = "L" + index,
            Amount = index
        }).ToList()
    };

    /// <summary>检查合同条件。</summary>
    private static void Require(bool condition, string scenario, string message)
    {
        if (!condition)
            throw new InvalidOperationException($"Entity PageHeader contract {scenario}: {message}");
    }

    /// <summary>保存合同单据。</summary>
    private sealed class HeaderOrder
    {
        /// <summary>获取或设置明细。</summary>
        public List<HeaderLine> Lines { get; set; } = new();

        /// <summary>获取或设置占用测试值。</summary>
        public string Collision { get; set; }
    }

    /// <summary>保存合同明细。</summary>
    private sealed class HeaderLine
    {
        /// <summary>获取或设置编码。</summary>
        public string Code { get; set; }

        /// <summary>获取或设置金额。</summary>
        public decimal Amount { get; set; }
    }

    /// <summary>读取单元格的探针。</summary>
    private sealed class CellProbe
    {
        /// <summary>获取或设置单元格文本。</summary>
        public string Value { get; set; }
    }

}
