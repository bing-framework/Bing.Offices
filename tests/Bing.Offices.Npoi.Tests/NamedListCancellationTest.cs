using System;
using System.Reflection;
using System.Threading;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证名称字典扫描响应导入取消令牌。
/// </summary>
public sealed class NamedListCancellationTest
{
    /// <summary>
    /// 验证读取首个字典行后取消会中断后续扫描。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ValidateCell_NamedListCanceledDuringScan_ShouldThrow(bool xls)
    {
        using IWorkbook source = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var data = source.CreateSheet("Data");
        var lookup = source.CreateSheet("Lookup");
        lookup.CreateRow(0).CreateCell(0).SetCellValue("Yes");
        lookup.CreateRow(1).CreateCell(0).SetCellValue("No");
        var name = source.CreateName();
        name.NameName = "Allowed";
        name.RefersToFormula = "Lookup!$A$1:$A$2";
        var helper = data.GetDataValidationHelper();
        var validation = helper.CreateValidation(helper.CreateFormulaListConstraint("Allowed"),
            new CellRangeAddressList(0, 0, 0, 0));

        using var cancellation = new CancellationTokenSource();
        var workbook = DispatchProxy.Create<IWorkbook, WorkbookProxy>();
        var dataSheet = DispatchProxy.Create<ISheet, SheetProxy>();
        var lookupSheet = DispatchProxy.Create<ISheet, SheetProxy>();
        ((WorkbookProxy)(object)workbook).Configure(source, data, dataSheet, lookupSheet);
        ((SheetProxy)(object)dataSheet).Configure(data, workbook, null);
        ((SheetProxy)(object)lookupSheet).Configure(lookup, workbook, cancellation);
        var value = new ExcelCellValue("Yes", "Yes", ExcelCellKind.Text);

        Assert.Throws<OperationCanceledException>(() =>
            NpoiWorkbookValidationPipeline.ValidateCell(validation, value, "Yes", dataSheet,
                false, cancellation.Token, out _));
        Assert.True(cancellation.IsCancellationRequested);
    }

    /// <summary>
    /// 转发工作簿访问并保留被测工作表代理。
    /// </summary>
    public class WorkbookProxy : DispatchProxy
    {
        /// <summary>
        /// 获取实际工作簿。
        /// </summary>
        private IWorkbook _source;

        /// <summary>
        /// 获取实际业务工作表。
        /// </summary>
        private ISheet _data;

        /// <summary>
        /// 获取业务工作表代理。
        /// </summary>
        private ISheet _dataProxy;

        /// <summary>
        /// 获取字典工作表代理。
        /// </summary>
        private ISheet _lookupProxy;

        /// <summary>
        /// 初始化代理所需的工作簿和工作表。
        /// </summary>
        public void Configure(IWorkbook source, ISheet data, ISheet dataProxy, ISheet lookupProxy)
        {
            _source = source;
            _data = data;
            _dataProxy = dataProxy;
            _lookupProxy = lookupProxy;
        }

        /// <inheritdoc />
        protected override object Invoke(MethodInfo method, object[] arguments)
        {
            if (method.Name == nameof(IWorkbook.GetSheetIndex)
                && ReferenceEquals(arguments[0], _dataProxy))
                return _source.GetSheetIndex(_data);
            if (method.Name == nameof(IWorkbook.GetSheet)
                && Equals(arguments[0], "Lookup"))
                return _lookupProxy;
            return method.Invoke(_source, arguments);
        }
    }

    /// <summary>
    /// 在读取字典首行后触发取消。
    /// </summary>
    public class SheetProxy : DispatchProxy
    {
        /// <summary>
        /// 获取实际工作表。
        /// </summary>
        private ISheet _source;

        /// <summary>
        /// 获取工作簿代理。
        /// </summary>
        private IWorkbook _workbook;

        /// <summary>
        /// 获取读取字典时使用的取消源。
        /// </summary>
        private CancellationTokenSource _cancellation;

        /// <summary>
        /// 初始化代理所需的工作表和取消源。
        /// </summary>
        public void Configure(ISheet source, IWorkbook workbook, CancellationTokenSource cancellation)
        {
            _source = source;
            _workbook = workbook;
            _cancellation = cancellation;
        }

        /// <inheritdoc />
        protected override object Invoke(MethodInfo method, object[] arguments)
        {
            if (method.Name == "get_Workbook")
                return _workbook;
            if (method.Name == nameof(ISheet.GetRow) && Equals(arguments[0], 0))
                _cancellation?.Cancel();
            return method.Invoke(_source, arguments);
        }
    }
}
