using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// NPOI Workbook 命名列表导入校验的职责级测试。
/// </summary>
public sealed class NamedListValidationTest
{
    /// <summary>
    /// 验证命名列表支持局部优先、全局回退以及绝对单格、单行和单列区域。
    /// </summary>
    /// <param name="xls">是否使用 HSSF XLS 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void NamedList_ShouldResolveScopesAndBoundedRanges(bool xls)
    {
        using var source = new MemoryStream(CreateWorkbook(xls, workbook =>
        {
            var data = workbook.CreateSheet("Data");
            var lookup = workbook.CreateSheet("Lookup");
            SetValues(lookup, 0, "Global", "GlobalAlt", "Single");
            SetValues(lookup, 1, "GlobalAlt", "Local", "RowB");
            SetValues(lookup, 2, null, null, null, "ColumnA");
            SetValues(lookup, 3, null, null, null, "ColumnB");
            SetValues(data, 0, "Value");
            SetValues(data, 1, "Local");
            SetValues(data, 2, "Global");
            SetValues(data, 3, "GlobalAlt");
            SetValues(data, 4, "Single");
            SetValues(data, 5, "RowB");
            SetValues(data, 6, "ColumnB");

            AddName(workbook, "Allowed", "'Lookup'!$A$1:$A$2");
            AddName(workbook, "Allowed", "'Lookup'!$B$2", workbook.GetSheetIndex(data));
            AddName(workbook, "GlobalOnly", "'Lookup'!$A$1:$A$2");
            AddName(workbook, "SingleCell", "'Lookup'!$C$1");
            AddName(workbook, "SingleRow", "'Lookup'!$B$2:$C$2");
            AddName(workbook, "SingleColumn", "'Lookup'!$D$3:$D$4");

            AddListValidation(data, 1, "Allowed");
            AddListValidation(data, 2, "Allowed");
            AddListValidation(data, 3, "GlobalOnly");
            AddListValidation(data, 4, "SingleCell");
            AddListValidation(data, 5, "SingleRow");
            AddListValidation(data, 6, "SingleColumn");
        }));
        var request = ExcelImport.Workbook<Root>(builder => builder
            .ValidationMode(ExcelImportValidationMode.WorkbookRules)
            .Sheet("Data", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

        var result = new NpoiExcelImporter().Import(source, request);

        Assert.Equal(new[] { "Local", "GlobalAlt", "Single", "RowB", "ColumnB" },
            result.Workbook.Rows.ConvertAll(item => item.Value));
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal(3, error.RowIndex);
        Assert.Contains("显式列表", error.Message);
    }

    /// <summary>
    /// 验证无效局部名称遮蔽有效全局名称，并报告结构化 Workbook 校验错误。
    /// </summary>
    /// <param name="xls">是否使用 HSSF XLS 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void NamedList_InvalidLocalName_ShouldNotFallBackToGlobal(bool xls)
    {
        using var source = new MemoryStream(CreateWorkbook(xls, workbook =>
        {
            var data = workbook.CreateSheet("Data");
            var lookup = workbook.CreateSheet("Lookup");
            SetValues(lookup, 0, "Allowed");
            SetValues(data, 0, "Value");
            SetValues(data, 1, "Allowed");
            AddName(workbook, "Broken", "'Lookup'!$A$1");
            AddName(workbook, "Broken", "OFFSET('Lookup'!$A$1,0,0,1,1)", workbook.GetSheetIndex(data));
            AddListValidation(data, 1, "Broken");
        }));
        var request = CreateRequest();

        var result = new NpoiExcelImporter().Import(source, request);

        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Contains("暂不支持", error.Message);
    }

    /// <summary>
    /// 验证公式、名称链、二维区域、外部引用和相对引用均作为 Unsupported 处理。
    /// </summary>
    /// <param name="xls">是否使用 HSSF XLS 工作簿。</param>
    /// <param name="refersTo">名称保存的引用公式。</param>
    [Theory]
    [InlineData(true, "OFFSET('Lookup'!$A$1,0,0,2,1)")]
    [InlineData(false, "OtherList")]
    [InlineData(false, "'Lookup'!$A$1:$B$2")]
    [InlineData(false, "'Lookup'!A1:A2")]
    [InlineData(true, "'Lookup'!$A$1:$A$2,'Lookup'!$C$1:$C$2")]
    public void NamedList_UnsupportedDefinitions_ShouldReturnStructuredError(bool xls, string refersTo)
    {
        using var source = new MemoryStream(CreateWorkbook(xls, workbook =>
        {
            var data = workbook.CreateSheet("Data");
            var lookup = workbook.CreateSheet("Lookup");
            SetValues(lookup, 0, "Allowed", "Other");
            SetValues(lookup, 1, "More", "Values");
            SetValues(data, 0, "Value");
            SetValues(data, 1, "Allowed");
            if (refersTo == "OtherList")
                AddName(workbook, "OtherList", "'Lookup'!$A$1");
            AddName(workbook, "BadList", refersTo);
            AddListValidation(data, 1, "BadList");
        }));
        var request = CreateRequest();

        var result = new NpoiExcelImporter().Import(source, request);

        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Contains("暂不支持", error.Message);
    }

    /// <summary>
    /// 验证 XLSX 名称中的外部工作簿引用返回结构化校验错误。
    /// </summary>
    [Fact]
    public void NamedList_ExternalReference_ShouldReturnStructuredError()
    {
        using var source = new MemoryStream(CreateExternalReferenceWorkbook());

        var result = new NpoiExcelImporter().Import(source, CreateRequest());

        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Contains("暂不支持", error.Message);
    }

    /// <summary>
    /// 验证字典公式单元格不会被当作文本允许值。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void NamedList_FormulaCell_ShouldReturnStructuredError(bool xls)
    {
        using var source = new MemoryStream(CreateWorkbook(xls, workbook =>
        {
            var data = workbook.CreateSheet("Data");
            var lookup = workbook.CreateSheet("Lookup");
            lookup.CreateRow(0).CreateCell(0).SetCellFormula("1+1");
            SetValues(data, 0, "Value");
            SetValues(data, 1, "2");
            AddName(workbook, "FormulaList", "Lookup!$A$1");
            AddListValidation(data, 1, "FormulaList");
        }));

        var result = new NpoiExcelImporter().Import(source, CreateRequest());
        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Contains("暂不支持", error.Message);
    }

    /// <summary>
    /// 验证名称来源在同一工作簿内复用且不同工作簿互不影响。
    /// </summary>
    /// <param name="xls">是否使用 HSSF XLS 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void NamedList_ShouldCacheSourcePerWorkbook(bool xls)
    {
        using IWorkbook first = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var data = first.CreateSheet("Data");
        var lookup = first.CreateSheet("Lookup");
        lookup.CreateRow(0).CreateCell(0).SetCellValue("Allowed");
        AddName(first, "AllowedList", "Lookup!$A$1");
        Assert.True(NpoiWorkbookValidationPipeline.IsSupportedNamedListSource(
            data, "AllowedList", CancellationToken.None));

        lookup.GetRow(0).GetCell(0).SetCellFormula("1+1");
        Assert.True(NpoiWorkbookValidationPipeline.IsSupportedNamedListSource(
            data, "AllowedList", CancellationToken.None));

        using IWorkbook second = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var secondData = second.CreateSheet("Data");
        var secondLookup = second.CreateSheet("Lookup");
        secondLookup.CreateRow(0).CreateCell(0).SetCellFormula("1+1");
        AddName(second, "AllowedList", "Lookup!$A$1");
        Assert.False(NpoiWorkbookValidationPipeline.IsSupportedNamedListSource(
            secondData, "AllowedList", CancellationToken.None));
    }

    /// <summary>
    /// 验证带特殊字符的合法工作表名称可用于名称列表。
    /// </summary>
    /// <param name="xls">是否使用 HSSF XLS 工作簿。</param>
    /// <param name="sheetName">字典工作表名称。</param>
    /// <param name="reference">名称引用表达式。</param>
    [Theory]
    [InlineData(true, "Lookup,2026", "'Lookup,2026'!$A$1")]
    [InlineData(false, "Lookup,2026", "'Lookup,2026'!$A$1")]
    [InlineData(false, "Lookup!(2026)", "'Lookup!(2026)'!$A$1")]
    [InlineData(false, "O'Brien", "'O''Brien'!$A$1")]
    public void NamedList_QuotedSheetNameWithSpecialCharacters_ShouldImport(
        bool xls, string sheetName, string reference)
    {
        using var source = new MemoryStream(CreateWorkbook(xls, workbook =>
        {
            var data = workbook.CreateSheet("Data");
            var lookup = workbook.CreateSheet(sheetName);
            SetValues(lookup, 0, "Allowed");
            SetValues(data, 0, "Value");
            SetValues(data, 1, "Allowed");
            AddName(workbook, "AllowedList", reference);
            AddListValidation(data, 1, "AllowedList");
        }));

        var result = new NpoiExcelImporter().Import(source, CreateRequest());
        Assert.Empty(result.Errors);
        Assert.Equal("Allowed", Assert.Single(result.Workbook.Rows).Value);
    }

    /// <summary>
    /// 创建名称列表工作簿导入请求。
    /// </summary>
    /// <returns>配置为工作簿原生校验模式的导入请求。</returns>
    private static ExcelWorkbookImportRequest<Root> CreateRequest() =>
        ExcelImport.Workbook<Root>(builder => builder
            .ValidationMode(ExcelImportValidationMode.WorkbookRules)
            .Sheet("Data", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

    /// <summary>
    /// 创建并序列化名称列表测试工作簿。
    /// </summary>
    /// <param name="xls">是否创建 XLS 工作簿。</param>
    /// <param name="configure">工作簿配置委托。</param>
    /// <returns>工作簿字节内容。</returns>
    private static byte[] CreateWorkbook(bool xls, Action<IWorkbook> configure)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        configure(workbook);
        using var output = new MemoryStream();
        workbook.Write(output, leaveOpen: true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建引用外部工作簿的名称列表测试文件。
    /// </summary>
    /// <returns>测试工作簿字节内容。</returns>
    private static byte[] CreateExternalReferenceWorkbook()
    {
        var bytes = CreateWorkbook(false, workbook =>
        {
            var data = workbook.CreateSheet("Data");
            var lookup = workbook.CreateSheet("Lookup");
            SetValues(lookup, 0, "Allowed", "Other");
            SetValues(lookup, 1, "More", "Values");
            SetValues(data, 0, "Value");
            SetValues(data, 1, "Allowed");
            AddName(workbook, "BadList", "'Lookup'!$A$1:$A$2");
            AddListValidation(data, 1, "BadList");
        });

        using var package = new MemoryStream();
        package.Write(bytes, 0, bytes.Length);
        package.Position = 0;
        using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/workbook.xml");
            XDocument document;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                document = XDocument.Load(reader);
            var name = document.Descendants().Single(element => element.Name.LocalName == "definedName"
                && (string)element.Attribute("name") == "BadList");
            name.Value = "'[External.xlsx]Lookup'!$A$1:$A$2";
            entry.Delete();
            using var output = archive.CreateEntry("xl/workbook.xml").Open();
            document.Save(output);
        }

        return package.ToArray();
    }

    /// <summary>
    /// 向测试工作簿添加名称定义。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="name">名称。</param>
    /// <param name="refersTo">名称引用公式。</param>
    /// <param name="sheetIndex">局部名称所属的工作表索引；负数表示全局名称。</param>
    private static void AddName(IWorkbook workbook, string name, string refersTo, int sheetIndex = -1)
    {
        var definedName = workbook.CreateName();
        if (sheetIndex >= 0)
            definedName.SheetIndex = sheetIndex;
        definedName.NameName = name;
        definedName.RefersToFormula = refersTo;
    }

    /// <summary>
    /// 为指定单元格添加名称列表校验。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="rowIndex">目标单元格的零基行号。</param>
    /// <param name="name">列表名称。</param>
    private static void AddListValidation(ISheet sheet, int rowIndex, string name)
    {
        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateFormulaListConstraint(name);
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(rowIndex, rowIndex, 0, 0)));
    }

    /// <summary>
    /// 写入测试工作表的一行文本值。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="rowIndex">目标行的零基行号。</param>
    /// <param name="values">待写入的文本值。</param>
    private static void SetValues(ISheet sheet, int rowIndex, params string[] values)
    {
        var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
        for (var index = 0; index < values.Length; index++)
        {
            if (values[index] != null)
                row.CreateCell(index).SetCellValue(values[index]);
        }
    }

    /// <summary>
    /// 名称列表导入根模型。
    /// </summary>
    private sealed class Root
    {
        /// <summary>
        /// 获取导入行集合。
        /// </summary>
        public List<Row> Rows { get; } = new List<Row>();
    }

    /// <summary>
    /// 名称列表导入测试行。
    /// </summary>
    private sealed class Row
    {
        /// <summary>
        /// 获取或设置列表值。
        /// </summary>
        public string Value { get; set; }
    }
}
