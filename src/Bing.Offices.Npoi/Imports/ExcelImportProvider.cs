using Bing.Extensions;
using Bing.Helpers;
using Bing.Offices.Attributes;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.Metadata.Excels;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Metadata.Excels;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Bing.Offices.Conversions;
using Bing.Offices.Npoi.Conversions;
using ICell = Bing.Offices.Metadata.Excels.ICell;
using IRow = Bing.Offices.Metadata.Excels.IRow;
using IWorkbook = Bing.Offices.Metadata.Excels.IWorkbook;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// Excel导入提供程序
/// </summary>
public class ExcelImportProvider : IExcelImportProvider
{
    /// <summary>
    /// 哈希表
    /// </summary>
    private static Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));

    /// <summary>
    /// 哈希表动态单元格
    /// </summary>
    private static Hashtable TableDynamicCell = Hashtable.Synchronized(new Hashtable(1024));

    /// <summary>
    /// 单元格值转换器
    /// </summary>
    private readonly ICellValueConverter _converter;

    /// <summary>
    /// 初始化一个<see cref="ExcelImportProvider"/>类型的实例
    /// </summary>
    /// <param name="converter">单元格值转换器</param>
    public ExcelImportProvider(ICellValueConverter converter = null) => _converter = converter ?? new CellValueConverter();

    /// <summary>
    /// 转换
    /// </summary>
    /// <typeparam name="TTemplate">导入模板类型</typeparam>
    /// <param name="fileUrl">文件地址</param>
    /// <param name="sheetIndex">工作表索引</param>
    /// <param name="headerRowIndex">标题行索引</param>
    /// <param name="dataRowStartIndex">数据行起始索引</param>
    /// <param name="multiSheet">是否支持多工作表模式</param>
    /// <param name="maxColumnLength">最大列长度</param>
    /// <param name="enabledEmptyLine">启用空行模式。启用时，行内遇到空行将抛出异常错误信息</param>
    public virtual IWorkbook Convert<TTemplate>(string fileUrl, int sheetIndex = 0, int headerRowIndex = 0, int dataRowStartIndex = 1, bool multiSheet = false, int maxColumnLength = 100, bool enabledEmptyLine = false) where TTemplate : class, new()
    {
        IImportOptions options = new ImportOptions
        {
            FileUrl = fileUrl,
            SheetIndex = sheetIndex,
            HeaderRowIndex = headerRowIndex,
            DataRowIndex = dataRowStartIndex,
            MultiSheet = multiSheet,
            MaxColumnLength = maxColumnLength,
            EnabledEmptyLine = enabledEmptyLine
        };
        return Convert<TTemplate>(options);
    }

    /// <summary>
    /// 转换
    /// </summary>
    /// <typeparam name="TTemplate">导入模板类型</typeparam>
    /// <param name="options">导入选项配置</param>
    public virtual IWorkbook Convert<TTemplate>(IImportOptions options) where TTemplate : class, new()
    {
        CleanHeaderRowCache(options);
        var workbook = new NpoiWorkbook();
        var innerWorkbook = GetWorkbook(options.FileUrl);
        if (options.MultiSheet == false)
        {
            BuildSheet<TTemplate>(workbook, innerWorkbook, options);
            return workbook;
        }

        for (var i = 0; i < innerWorkbook.NumberOfSheets; i++)
            BuildSheet<TTemplate>(workbook, innerWorkbook, options);
        return workbook;
    }

    /// <summary>
    /// 清理缓存
    /// </summary>
    /// <param name="options">导入选项配置</param>
    private void CleanHeaderRowCache(IImportOptions options)
    {
        if (!options.EnabledHeaderRowCache)
        {
            Table = Hashtable.Synchronized(new Hashtable(1024));
            TableDynamicCell = Hashtable.Synchronized(new Hashtable(1024));
        }
    }

    /// <summary>
    /// 构建工作表
    /// </summary>
    /// <typeparam name="TTemplate">导入模板类型</typeparam>
    /// <param name="workbook">工作簿</param>
    /// <param name="innerWorkbook">内部工作簿</param>
    /// <param name="options">导入选项配置</param>
    private void BuildSheet<TTemplate>(IWorkbook workbook, NPOI.SS.UserModel.IWorkbook innerWorkbook, IImportOptions options)
    {
        var innerSheet = GetSheet(innerWorkbook, options.SheetIndex);
        //if (innerSheet.GetRow(0).PhysicalNumberOfCells > maxColumnLength)
        //    throw new OfficeException($"导入数据初始化过多的无效列: {maxColumnLength}");
        var sheet = workbook.CreateSheet(innerSheet.SheetName, options.HeaderRowIndex);
        HandleHeader(sheet, innerSheet, options.HeaderRowIndex);
        VerifyHeader<TTemplate>(sheet, options);
        HandleBody<TTemplate>(sheet, innerSheet, options);
    }
        
    /// <summary>
    /// 获取工作簿
    /// </summary>
    /// <param name="fileUrl">文件地址</param>
    private NPOI.SS.UserModel.IWorkbook GetWorkbook(string fileUrl)
    {
        if (string.IsNullOrWhiteSpace(fileUrl))
            throw new ArgumentNullException(nameof(fileUrl));
        if (!File.Exists(fileUrl))
            throw new FileNotFoundException("找不到文件", fileUrl);
        var ext = Path.GetExtension(fileUrl).ToLower().Trim();
        if (ext != ".xls" && ext != ".xlsx")
            throw new NotSupportedException("仅支持后缀名为.xls或者.xlsx的文件");
        return WorkbookFactory.Create(fileUrl);
    }

    /// <summary>
    /// 获取工作表
    /// </summary>
    /// <param name="workbook">工作簿</param>
    /// <param name="sheetIndex">工作表索引</param>
    private ISheet GetSheet(NPOI.SS.UserModel.IWorkbook workbook, int sheetIndex = 0) => workbook.GetSheetAt(sheetIndex);

    /// <summary>
    /// 处理表头
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="innerSheet">NPOI工作表</param>
    /// <param name="headerRowIndex">标题行索引</param>
    private void HandleHeader(IWorkSheet sheet, ISheet innerSheet, int headerRowIndex)
    {
        var innerRow = innerSheet.GetRow(headerRowIndex);
        var cells = new List<ICell>();
        foreach (var cell in innerRow.Cells.Where(x => !string.IsNullOrEmpty(GetStringValue(x))))
            cells.Add(new Cell(GetStringValue(cell)) {ColumnIndex = cell.ColumnIndex, Name = GetStringValue(cell) });
        sheet.AddHeadRow(cells.ToArray());
    }

    /// <summary>
    /// 验证表头
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="options">导入选项配置</param>
    private void VerifyHeader<TTemplate>(IWorkSheet sheet, IImportOptions options)
    {
        var header = sheet.GetHeader().LastOrDefault();
        if (header == null)
            throw new OfficeHeaderException($"导入的模板不正确，未匹配表头", options.HeaderRowIndex);

        var cellNames = header.Cells.GroupBy(x => x.Name).Select(x => new {Name = x.Key, Count = x.Count()}).ToList();
        if (options.HeaderColumnOnly && cellNames.Any(x => x.Count > 1))
        {
            throw new OfficeHeaderException($"导入的表格存在重复列:{cellNames.Where(x => x.Count > 1).Select(x => x.Name).ExpandAndToString()}", options.HeaderRowIndex);
        }

        if (!options.HeaderMatch) return;
        List<PropertyInfo> properties = typeof(TTemplate).GetProperties().ToList();
        var props = properties
            .Where(x => x.GetCustomAttribute<DynamicColumnAttribute>() == null)
            .Where(x => x.GetCustomAttribute<ExcelIgnoreAttribute>() == null)
            .Select(p => new
            {
                Name = p.GetCustomAttribute<ColumnNameAttribute>()?.Name,
                Code = p.Name
            }).ToList();

        var cellName = cellNames.Select(x => x.Name);
        var list = props.Where(p => !cellName.Contains(p.Name)).ToList();
        if (list.Any())
        {
            list = list.Where(p => !(string.IsNullOrWhiteSpace(p.Name) && cellName.Contains(p.Code))).ToList();
            if (list.Any())
                throw new OfficeHeaderException($"导入的表格不存在列：{list.Select(x => string.IsNullOrWhiteSpace(x.Name) ? x.Code : x.Name).ExpandAndToString()}", options.HeaderRowIndex);
        }
        else if (options.MappingDictionary != null && options.MappingDictionary.Any())
        {
            var dic = options.MappingDictionary.ToList().Where(x => !cellName.Contains(x.Value)).ToList();
            if (dic.Count > 0)
                throw new OfficeHeaderException($"导入的表格不存在列：{dic.Select(x => x.Value).ExpandAndToString()}", options.HeaderRowIndex);
        }
    }

    /// <summary>
    /// 处理正文
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="innerSheet">NPOI工作表</param>
    /// <param name="options">导入选项配置</param>
    private void HandleBody<TTemplate>(IWorkSheet sheet, ISheet innerSheet, IImportOptions options)
    {
        var header = sheet.GetHeader().LastOrDefault();
        // LastRowNum: 获取最后一行的行数，如果sheet中一行数据都没有则返回-1，只有第一行有数据则返回0，最后有数据的行是第n行则返回n-1
        // PhysicalNumberOfRows: 获取有记录的行数，即：最后有数据的行是第n行，前面有m行是空行没数据，则返回n-m
        for (var i = options.DataRowIndex; i < innerSheet.GetHasDataRowNum() + 1; i++)
        {
            var innerRow = innerSheet.GetRow(i);
            if (CheckIgnoreEmptyLine(options, innerRow.IsEmptyRow())) break;
            if (CheckEmptyLine(innerRow.IsEmptyRow(), options.EnabledEmptyLine, i)) continue;
            sheet.AddBodyRow(Convert<TTemplate>(innerRow, header), innerRow.RowNum);
        }
    }

    /// <summary>
    /// 忽略空行后数据
    /// </summary>
    /// <param name="options">导入选项配置</param>
    /// <param name="isEmptyRow">是否空行</param>
    private bool CheckIgnoreEmptyLine(IImportOptions options, bool isEmptyRow) => options.IgnoreEmptyLineAfterData && isEmptyRow;

    /// <summary>
    /// 检查空行
    /// </summary>
    /// <param name="isEmptyRow">是否空行</param>
    /// <param name="enabledEmptyLine">启用空行模式。启用时，行内遇到空行将抛出异常错误信息</param>
    /// <param name="rowIndex">行号</param>
    private bool CheckEmptyLine(bool isEmptyRow, bool enabledEmptyLine,int rowIndex)
    {
        if (!enabledEmptyLine)
            return isEmptyRow;
        if (isEmptyRow)
            throw new OfficeEmptyLineException($"导入数据存在空行", rowIndex);
        return false;
    }

    /// <summary>
    /// 转换
    /// </summary>
    /// <typeparam name="TTemplate">模板类型</typeparam>
    /// <param name="row">NPOI单元行</param>
    /// <param name="header">表头行</param>
    private IList<ICell> Convert<TTemplate>(NPOI.SS.UserModel.IRow row, IRow header)
    {
        var type = typeof(TTemplate);
        var props = type.GetProperties().ToList();
        var cells = new List<ICell>();
        var hash = string.Empty;
        if (type.GetCustomAttribute<HasDynamicColumnAttribute>() != null)
            hash = $"_{Encrypt.Md5By32(header.Cells.Select(x => x.Name).Join())}";

        string columnName;
        foreach (var cell in header.Cells)
        {
            columnName = cell?.Name;
            if (string.IsNullOrWhiteSpace(columnName))
                continue;
            var key = $"{type.FullName}_{hash}_{cell.ColumnIndex}";

            if (Table[key] == null)
            {
                var isDynamic = false;
                // 优先匹配列名特性值
                var matchProperty = props.FirstOrDefault(x =>
                    x.GetCustomAttribute<ColumnNameAttribute>()?.Name == columnName);
                // 次之匹配属性名
                if (matchProperty == null)
                    matchProperty = props.FirstOrDefault(x =>
                        x.Name.Equals(columnName, StringComparison.CurrentCultureIgnoreCase));
                // 次之匹配动态列
                if (matchProperty == null)
                {
                    matchProperty = props.FirstOrDefault(x => x.GetCustomAttribute<DynamicColumnAttribute>() != null);
                    if (matchProperty != null)
                        isDynamic = true;
                }
                var propertyName = matchProperty?.Name;
                Table[key] = propertyName;
                TableDynamicCell[key] = isDynamic;
            }

            var value = row.GetCell(cell.ColumnIndex) == null
                ? string.Empty
                : GetStringValue(row.GetCell(cell.ColumnIndex));
            cells.Add(new Cell(value)
            {
                ColumnIndex = cell.ColumnIndex,
                Name = columnName,
                PropertyName = Table[key]?.ToString(),
                IsDynamic = Bing.Helpers.Conv.ToBool(TableDynamicCell[key])
            });
        }
        return cells;
    }

    /// <summary>
    /// 获取单元格值字符串
    /// </summary>
    /// <param name="cell">单元格</param>
    protected virtual string GetStringValue(NPOI.SS.UserModel.ICell cell) => _converter.GetStringValue(cell);
}