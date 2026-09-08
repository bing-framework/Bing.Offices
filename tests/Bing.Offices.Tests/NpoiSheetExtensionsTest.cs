using System;
using System.Collections.Generic;
using System.Reflection;
using Bing.Offices.Metadata;
using Bing.Offices.Npoi.Extensions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// SheetExtensions 的行操作和合并区域职责测试。
/// </summary>
public sealed class NpoiSheetExtensionsTest
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RowQueries_ShouldHandleEmptyAndSparseRows(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Rows");

        var emptyRegions = SheetExtensions.GetAllMergedRegions(sheet);

        Assert.Equal(0, emptyRegions.GetLength(0));
        Assert.Equal(4, emptyRegions.GetLength(1));
        Assert.Equal(-1, SheetExtensions.GetHasDataRowNum(sheet));

        sheet.CreateRow(0).CreateCell(0).SetCellValue("first");
        sheet.CreateRow(3).CreateCell(0).SetCellValue("third");
        sheet.CreateRow(6);

        Assert.Equal(3, SheetExtensions.GetHasDataRowNum(sheet));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AddRow_ShouldCreateRowsFromOneAndLeaveEmptyInputUntouched(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Rows");

        SheetExtensions.AddRow(sheet, new[] { "one", "two" }, (row, value) =>
        {
            row.CreateCell(0).SetCellValue(value);
        });

        var lastRowBeforeEmptyInput = sheet.LastRowNum;
        SheetExtensions.AddRow(sheet, Array.Empty<string>(), (_, _) =>
        {
            throw new InvalidOperationException("空集合不应执行行回调");
        });

        Assert.Equal(1, sheet.GetRow(1).RowNum);
        Assert.Equal("one", sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal(2, sheet.GetRow(2).RowNum);
        Assert.Equal("two", sheet.GetRow(2).GetCell(0).StringCellValue);
        Assert.Equal(400, sheet.GetRow(1).Height);
        Assert.Equal(400, sheet.GetRow(2).Height);
        Assert.Equal(lastRowBeforeEmptyInput, sheet.LastRowNum);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InsertRowAndRows_ShouldShiftDataAndMergedRegions(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Rows");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("zero");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("one");
        sheet.CreateRow(2).CreateCell(0).SetCellValue("two");
        sheet.CreateRow(3);
        SheetExtensions.AddMergedRegion(sheet, new MergedRegionInfo(0, 2, 3, 0, 1));

        var insertedRow = SheetExtensions.InsertRow(sheet, 1);

        Assert.NotNull(insertedRow);
        Assert.Same(insertedRow, sheet.GetRow(1));
        Assert.Equal("one", sheet.GetRow(2).GetCell(0).StringCellValue);
        Assert.Equal("two", sheet.GetRow(3).GetCell(0).StringCellValue);
        AssertMergedBounds(sheet, new[] { 3, 4, 0, 1 });

        var oldLastRow = sheet.LastRowNum;
        var appendedRows = SheetExtensions.InsertRows(sheet, oldLastRow + 1, 2);

        Assert.Equal(2, appendedRows.Length);
        Assert.Equal(oldLastRow + 1, appendedRows[0].RowNum);
        Assert.Equal(oldLastRow + 2, appendedRows[1].RowNum);
        Assert.Null(appendedRows[0].GetCell(0));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RemoveRows_ShouldHandleSparseRowsAndShiftFollowingMergedRegions(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Rows");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("first");
        sheet.CreateRow(3).CreateCell(0).SetCellValue("following");
        sheet.CreateRow(4);
        SheetExtensions.AddMergedRegion(sheet, new MergedRegionInfo(0, 3, 4, 0, 1));

        var removedCount = SheetExtensions.RemoveRows(sheet, 1, 2);

        Assert.Equal(2, removedCount);
        Assert.Equal("first", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("following", sheet.GetRow(1).GetCell(0).StringCellValue);
        AssertMergedBounds(sheet, new[] { 1, 2, 0, 1 });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RemoveRow_ShouldRemoveOneRowAndReturnOneAtAnEmptyBoundary(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Rows");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("kept");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("removed");

        Assert.Equal(1, SheetExtensions.RemoveRow(sheet, 1));
        Assert.Null(sheet.GetRow(1));
        Assert.Equal("kept", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal(1, SheetExtensions.RemoveRow(sheet, 50));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RowOperations_ShouldRejectNegativeAndEmptyRangesWithoutMutation(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Rows");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("unchanged");

        Assert.Equal("deleteRowStartIndex", Assert.Throws<ArgumentOutOfRangeException>(() =>
            SheetExtensions.DeleteRows(sheet, -1, 1)).ParamName);
        Assert.Equal("count", Assert.Throws<ArgumentOutOfRangeException>(() =>
            SheetExtensions.DeleteRows(sheet, 0, 0)).ParamName);
        Assert.Equal("rowIndex", Assert.Throws<ArgumentOutOfRangeException>(() =>
            SheetExtensions.InsertRow(sheet, -1)).ParamName);
        Assert.Equal("rowIndex", Assert.Throws<ArgumentOutOfRangeException>(() =>
            SheetExtensions.InsertRows(sheet, -1, 1)).ParamName);
        Assert.Equal("rowsCount", Assert.Throws<ArgumentOutOfRangeException>(() =>
            SheetExtensions.InsertRows(sheet, 0, 0)).ParamName);
        Assert.Equal("startRowIndex", Assert.Throws<ArgumentOutOfRangeException>(() =>
            SheetExtensions.RemoveRow(sheet, -1)).ParamName);
        Assert.Equal("startRowIndex", Assert.Throws<ArgumentOutOfRangeException>(() =>
            SheetExtensions.RemoveRows(sheet, -1, 0)).ParamName);
        Assert.Equal("endRowIndex", Assert.Throws<ArgumentOutOfRangeException>(() =>
            SheetExtensions.RemoveRows(sheet, 1, 0)).ParamName);

        Assert.Equal(0, sheet.LastRowNum);
        Assert.Equal("unchanged", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal(0, sheet.NumMergedRegions);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DeleteRows_ShouldShiftSparseDataAndMergedRegions(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Rows");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("first");
        sheet.CreateRow(3).CreateCell(0).SetCellValue("following");
        sheet.CreateRow(4);
        SheetExtensions.AddMergedRegion(sheet, new MergedRegionInfo(0, 3, 4, 0, 1));

        SheetExtensions.DeleteRows(sheet, 1, 2);

        Assert.Equal("first", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("following", sheet.GetRow(1).GetCell(0).StringCellValue);
        AssertMergedBounds(sheet, new[] { 1, 2, 0, 1 });

        var lastRowBeforeBoundaryDelete = sheet.LastRowNum;
        SheetExtensions.DeleteRows(sheet, lastRowBeforeBoundaryDelete + 10, 1);
        Assert.Equal(lastRowBeforeBoundaryDelete, sheet.LastRowNum);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AddMergedRegionAndGetAllMergedRegions_ShouldPreserveBoundsAndOrder(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Merged");

        SheetExtensions.AddMergedRegion(sheet, new MergedRegionInfo(99, 1, 2, 3, 4));
        SheetExtensions.AddMergedRegion(sheet, new MergedRegionInfo(100, 5, 6, 0, 2));

        var regions = SheetExtensions.GetAllMergedRegions(sheet);

        Assert.Equal(2, regions.GetLength(0));
        Assert.Equal(4, regions.GetLength(1));
        Assert.Equal(1, regions[0, 0]);
        Assert.Equal(3, regions[0, 1]);
        Assert.Equal(2, regions[0, 2]);
        Assert.Equal(4, regions[0, 3]);
        Assert.Equal(5, regions[1, 0]);
        Assert.Equal(0, regions[1, 1]);
        Assert.Equal(6, regions[1, 2]);
        Assert.Equal(2, regions[1, 3]);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GetMergedRegionInfos_ShouldSupportAllRegionsIntersectionAndOnlyInternal(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Merged");
        SheetExtensions.AddMergedRegion(sheet, new MergedRegionInfo(0, 1, 2, 1, 2));
        SheetExtensions.AddMergedRegion(sheet, new MergedRegionInfo(1, 3, 5, 3, 4));

        var all = SheetExtensions.GetMergedRegionInfos(sheet);
        var internalOnly = SheetExtensions.GetMergedRegionInfos(sheet, 0, 4, 0, 4, true);
        var intersecting = SheetExtensions.GetMergedRegionInfos(sheet, 0, 4, 0, 4, false);
        var noMatch = SheetExtensions.GetMergedRegionInfos(sheet, 10, 12, 10, 12, false);

        Assert.Equal(2, all.Count);
        Assert.Equal(0, all[0].Index);
        Assert.Equal(1, all[0].FirstRow);
        Assert.Equal(2, all[0].LastRow);
        var internalRegion = Assert.Single(internalOnly);
        Assert.Equal(0, internalRegion.Index);
        Assert.Collection(intersecting,
            _ => { },
            region => Assert.Equal(3, region.FirstRow));
        Assert.Empty(noMatch);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RemoveMergedRegions_ShouldRemoveLocalRegionsThenAllRemainingRegions(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Merged");
        SheetExtensions.AddMergedRegion(sheet, new MergedRegionInfo(0, 0, 1, 0, 1));
        SheetExtensions.AddMergedRegion(sheet, new MergedRegionInfo(1, 3, 4, 0, 1));
        SheetExtensions.AddMergedRegion(sheet, new MergedRegionInfo(2, 6, 7, 0, 1));

        SheetExtensions.RemoveMergedRegions(sheet, 3, 4, 0, 1);

        Assert.Equal(2, sheet.NumMergedRegions);
        AssertMergedBounds(sheet, new[] { 0, 1, 0, 1 }, new[] { 6, 7, 0, 1 });

        SheetExtensions.RemoveMergedRegions(sheet);

        Assert.Equal(0, sheet.NumMergedRegions);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MoveMergedRegions_ShouldKeepCurrentProviderBehavior(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Merged");
        SheetExtensions.AddMergedRegion(sheet, new MergedRegionInfo(0, 2, 3, 2, 3));
        SheetExtensions.AddMergedRegion(sheet, new MergedRegionInfo(1, 7, 7, 6, 7));

        SheetExtensions.MoveMergedRegions(sheet, 1, 2);

        if (isXlsx)
            AssertMergedBounds(sheet, new[] { 2, 3, 2, 3 }, new[] { 7, 7, 6, 7 });
        else
            AssertMergedBounds(sheet, new[] { 3, 4, 4, 5 }, new[] { 8, 8, 8, 9 });

        SheetExtensions.MoveMergedRegions(sheet, 3, 4, 4, 5, -1, -1);

        if (isXlsx)
            AssertMergedBounds(sheet, new[] { 2, 3, 2, 3 }, new[] { 7, 7, 6, 7 });
        else
            AssertMergedBounds(sheet, new[] { 2, 3, 3, 4 }, new[] { 8, 8, 8, 9 });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MoveMergedRegions_ShouldRejectInvalidCoordinatesAndOverlapWithoutMutation(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Merged");
        SheetExtensions.AddMergedRegion(sheet, new MergedRegionInfo(0, 0, 1, 0, 1));
        SheetExtensions.AddMergedRegion(sheet, new MergedRegionInfo(1, 3, 4, 0, 1));

        Assert.Equal("moveRowCount", Assert.Throws<ArgumentOutOfRangeException>(() =>
            SheetExtensions.MoveMergedRegions(sheet, -1, 0)).ParamName);
        AssertMergedBounds(sheet, new[] { 0, 1, 0, 1 }, new[] { 3, 4, 0, 1 });

        Assert.Equal("moveColCount", Assert.Throws<ArgumentOutOfRangeException>(() =>
            SheetExtensions.MoveMergedRegions(sheet, 0, -1)).ParamName);
        AssertMergedBounds(sheet, new[] { 0, 1, 0, 1 }, new[] { 3, 4, 0, 1 });

        Assert.Equal("moveRowCount", Assert.Throws<ArgumentOutOfRangeException>(() =>
            SheetExtensions.MoveMergedRegions(sheet, 0, 1, 0, 1, int.MaxValue)).ParamName);
        AssertMergedBounds(sheet, new[] { 0, 1, 0, 1 }, new[] { 3, 4, 0, 1 });

        var overlapException = Assert.Throws<ArgumentException>(() =>
            SheetExtensions.MoveMergedRegions(sheet, 0, 1, 0, 1, 3));

        Assert.Equal("移动后的合并区域不能与现有区域重叠", overlapException.Message);
        Assert.Null(overlapException.ParamName);
        AssertMergedBounds(sheet, new[] { 0, 1, 0, 1 }, new[] { 3, 4, 0, 1 });
    }

    [Fact]
    public void RowAndMergeExtensions_UnknownSheet_ShouldKeepCurrentNoOpBehavior()
    {
        var unknownSheet = DispatchProxy.Create<ISheet, UnknownSheetProxy>();

        var merged = SheetExtensions.GetAllMergedRegions(unknownSheet);
        var infos = SheetExtensions.GetMergedRegionInfos(unknownSheet);

        Assert.Equal(0, merged.GetLength(0));
        Assert.Equal(4, merged.GetLength(1));
        Assert.Equal(-1, SheetExtensions.GetHasDataRowNum(unknownSheet));
        Assert.Empty(infos);
        SheetExtensions.RemoveMergedRegions(unknownSheet);
        SheetExtensions.MoveMergedRegions(unknownSheet, 1);
    }

    private static IWorkbook CreateWorkbook(bool isXlsx) =>
        isXlsx ? new XSSFWorkbook() : new HSSFWorkbook();

    private static void AssertMergedBounds(ISheet sheet, params int[][] expected)
    {
        Assert.Equal(expected.Length, sheet.NumMergedRegions);
        for (var i = 0; i < expected.Length; i++)
        {
            var actual = sheet.GetMergedRegion(i);
            Assert.Equal(expected[i][0], actual.FirstRow);
            Assert.Equal(expected[i][1], actual.LastRow);
            Assert.Equal(expected[i][2], actual.FirstColumn);
            Assert.Equal(expected[i][3], actual.LastColumn);
        }
    }

    private class UnknownSheetProxy : DispatchProxy
    {
        protected override object Invoke(MethodInfo targetMethod, object[] args) =>
            targetMethod.ReturnType.IsValueType ? Activator.CreateInstance(targetMethod.ReturnType) : null;
    }
}
