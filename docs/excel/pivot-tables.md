# Workbook 透视表

Workbook 导出可从已声明 Sheet 的有界数据区域创建原生透视表。ClosedXML XLSX 的基础模式已完成验收；NPOI XLSX 已接入基础及高级新建、完整来源缓存。职责、文件资源、双 TFM 包消费者及 Release/常规门禁结果属于较早候选的历史验证证据；2026-10-11 后续源码变化后的完整 Release、API 和包身份门禁尚未验证，当前执行记录有 `MSB3491 Access Denied`。高级模式在真实 Excel 中也复现文本 Count 为空，客户端完整验收仍未完成。NPOI XLS/HSSF、MiniExcel 和 SpreadCheetah 对透视请求在写入前返回 `UnsupportedFeature`。基础能力位为 `ExcelProviderFeatures.PivotTables`，高级模式另声明 `AdvancedPivotTables`；NPOI XLSX 与 ClosedXML XLSX 对已有 XLSX 模板保留/更新另声明 `TemplatePivotTables`。

## Provider 能力矩阵

| Provider / 格式 | 新建基础透视 | 新建高级透视 | 保留已有模板透视 | 更新已有模板透视 |
| --- | --- | --- | --- | --- |
| NPOI XLSX | 支持 | 支持；完整客户端数值验收未通过 | 支持 | 支持 |
| ClosedXML XLSX | 支持 | 支持；完整客户端数值验收未通过 | 支持 | 支持 |
| NPOI XLS/HSSF | 不支持，返回 `UnsupportedFeature` | 不支持，返回 `UnsupportedFeature` | 不支持 | 不支持 |
| MiniExcel、SpreadCheetah | 不支持透视导出 | 不支持透视导出 | 不支持 | 不支持 |
| ExcelDataReader | 仅导入，不提供透视导出 | 仅导入，不提供透视导出 | 不支持 | 不支持 |

模板保留和更新仅适用于 XLSX。`TemplatePivotTables` 表示 Provider 能保留或更新模板中的透视部件，不表示客户端已刷新或正确呈现结果。

| 行为 | 库侧语义 |
| --- | --- |
| 保留 | 保留模板定义、缓存和关系；不会根据新来源数据刷新缓存。 |
| 更新 | 重建本次更新透视的来源缓存和定义；新缓存写入 `refreshOnLoad=true`，仅请求客户端打开时刷新。 |
| 公式计算 | Provider 不计算公式；调用方须在导出前提供已计算的来源值。 |
| 刷新与呈现 | Provider 不运行客户端刷新或渲染。模板更新的客户端首开、刷新和保存重开尚未验收；高级模式的部分聚合数值也未通过最终验收。 |
| 数据权限 | 来源行及完整缓存仍保存在工作簿中；筛选、隐藏工作表或工作表保护不能作为敏感数据访问控制。 |

```csharp
var request = ExcelExport.Workbook(workbook => workbook
    .AddSheet("Data", rows)
    .PivotTable(new ExcelPivotTableDefinition
    {
        Name = "SalesByRegion",
        SourceSheet = "Data",
        SourceRange = new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0,
            EndRow = rows.Count, EndColumn = 1
        },
        TargetSheet = "Summary",
        TargetRow = 1, TargetColumn = 2,
        RowField = "Region",
        ValueField = "Amount"
    }));

using var output = new MemoryStream();
new ClosedXmlExcelExporter().Export(request, output);
```

坐标均从零开始，来源区域包含一行表头和至少一行数据；示例中的 `EndRow` 等于数据条数。表头须非空、唯一且与 `RowField`、`ValueField` 完全匹配（比较时忽略大小写，不自动修剪标题）。求和列的非空值必须是数值，且至少有一个数值；数字文本、日期和布尔值不能隐式参与求和。空金额忽略，空分组键允许，完整空数据行拒绝。目标 Sheet 会追加到工作簿，不能与现有业务 Sheet、其他透视目标或模板 Sheet 重名。构建请求时会快照透视定义及来源区域。

来源区域接受字面值，区域内任何字段的公式和 Excel 错误单元格都会在取值与缓存创建之前返回 `ConfigurationInvalid`。表头、分组、求和及未选择的缓存字段执行相同规则；模板公式有无缓存、缓存是否过期均不会改变规则。来源错误携带工作表和原始一基行列，消息包含透视名称。区域外公式继续按原有 Workbook 保存契约处理。

字段名称缺失时没有对应的来源单元格，错误定位到来源区域首个表头，并在消息中说明所需字段。空白或重复表头则定位到实际违规表头。标题前后空格保留；配置字段必须与保留空格后的完整标题匹配。

需要透视公式结果时，先在应用或显式公式计算接口中计算，再把结果作为数值写入来源。早期实现读取引擎 `Value` 时可能隐式求值，本次统一为明确的无公式来源合同。不会自动使用可能过期的公式缓存。Excel 数值和透视缓存使用引擎的数值存储，金额精度及舍入由应用定义。

来源名单只包含 `AddSheet` 声明的业务工作表，不能将前一个透视目标作为后续来源。目标名称在构建时检查 31 字符限制、非法字符及首尾单引号，失败不会枚举业务数据。来源区域与表格范围完全相同时，底层缓存可以通过表格名称引用来源；验收应解析名称对应的工作表和完整范围。

同时使用 `LookupList` 或 `CascadeList` 时，自动辅助 Sheet 会避开全部透视目标名称及模板已有 Sheet，按编号顺延。名称比较忽略大小写；合法的透视目标名称无需避开 `__BingLookup_` 或 `__BingCascade_` 前缀。字典名称引用与级联校验公式指向顺延后实际生成的辅助 Sheet。

当前单行字段求和布局预留两列，以及一行表头、实际分组行和一行总计。分组按单元格类型和值计数，文本比较忽略大小写，空分组键也占一行。目标起点、完整展开宽度和行数必须均在 XLSX 范围内；超限会在序列化前返回配置错误。这是当前布局的空间预算，客户端刷新或重新配置后的布局应由应用再次核对。

透视定义与缓存随 XLSX 写入；打开工作簿后的显示和刷新行为取决于 Excel 等客户端。本入口不计算公式、不在服务端执行客户端缓存刷新。显式调用 `PreserveTemplatePivotTables()` 时，NPOI XLSX 与 ClosedXML XLSX 会保留模板内已有透视；带业务 Sheet 的请求通过部件桥接写回原透视关系。`UpdatePivotTable(...)` 支持更新全部或部分模板透视；每个更新通过临时透视生成新缓存后替换原透视关系，未更新透视及其原始缓存继续保留。被更新透视共享缓存时会拆分新缓存，未更新消费者继续引用原缓存。单个请求可以同时使用 `UpdatePivotTable(...)` 更新模板透视并使用 `PivotTable(...)` 添加新透视；名称重复会在构建阶段拒绝。旧 `RowField`/`ValueField` 继续表示单行字段、单求和字段；新增高级集合进入独立配置路径，当前仍在完成高级客户端验收。导出到文件时仍使用原有原子提交；调用方传入的流仍归调用方所有。

更新来源数据时，可通过 `UseTemplate(...)` 指定已有模板，并用同名业务 Sheet 替换来源数据。更新定义仍须按原透视名称和目标工作表定位；未指定更新的透视保持原样：

```csharp
using var template = File.OpenRead("sales-template.xlsx");
var request = ExcelExport.Workbook(workbook => workbook
    .UseTemplate(template, leaveOpen: true)
    .AddSheet("Sales", refreshedRows)
    .UpdatePivotTable(new ExcelPivotTableDefinition
    {
        Name = "SalesByRegion",
        SourceSheet = "Sales",
        SourceRange = new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0,
            EndRow = refreshedRows.Count, EndColumn = 1
        },
        TargetSheet = "Summary",
        TargetRow = 1, TargetColumn = 2,
        RowField = "Region",
        ValueField = "Amount"
    }));

exporter.Export(request, output);
```

更新子集只重建来源缓存和目标透视部件，并为本次生成的缓存写入 `refreshOnLoad=true` 打开时刷新请求；Provider 不执行客户端刷新，也不保证 Excel、WPS 或 LibreOffice 打开时立即重新计算显示结果。需要确认最终汇总时仍须在目标客户端中刷新并检查完整结果。

模板更新还会检查透视和缓存部件的依赖安全：未知 `ext` 扩展、命名来源或 Provider 未声明的关系无法证明字段索引仍然有效时，更新在业务数据枚举前返回配置错误；显式保留模式继续原样保留这些部件。ClosedXML 自身已知的 `x14:pivotTableDefinition` 扩展按固定形状识别。清单通过图表部件的 `pivotSource` 识别透视图，并通过工作簿关系识别透视切片器缓存和时间线缓存；存在这些关联功能时，当前更新路径会在枚举业务数据前按功能名称拒绝更新，保留模式仍原样复制关系和部件。普通图表和未关联透视表的表格切片器不会触发该更新限制。透视图使用的 `pivotSource`、切片器缓存关系和时间线缓存关系分别由 [MS-ODRAWXML](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/9c323dc2-ee7d-43a3-85ef-b3bd7f44a78d) 与 [MS-XLSX](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/7dbb4481-b021-45cc-8bd4-6094b566a5ff)、[MS-XLSX Timeline Cache](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/29a0f58c-d942-4641-8ed0-4f02010326f2) 描述。该边界不等同于支持外部连接或任意 XML 扩展的更新。

## 高级模式与当前验收边界

`RowFields`、`ColumnFields`、`ValueFields` 和 `Filters` 可描述多层轴、多度量及报表筛选，能力位为 `AdvancedPivotTables`。显式高级集合不能与旧字段混用；至少一个行或列字段和一个值字段。同一来源值字段可参与多个聚合，但度量 `Caption` 须唯一，轴、筛选和值角色不得重复使用同一字段。

支持的配置聚合为 Sum、Count、Average、Minimum、Maximum。数值聚合跳过空值，拒绝数字文本、日期及布尔值，每个度量至少须有一个数值。Count 的合同是统计非空 Excel 单元格；显式空字符串缓存为文本，与真正空单元格不同。金额和大整数使用 Excel 数值存储，舍入与精度由应用明确处理。

筛选值接受文本、布尔、有限数值、DateTime 和代表空项的 null，按 Excel 类型匹配；文本忽略大小写且保留前后空白。筛选不删除来源行或缓存记录。源数据仍会完整保存在文件中，报表筛选不能用于隐藏敏感数据。Ascending/Descending 表示客户端标签排序，Excel 自定义列表及地区设置可能影响实际顺序。

`MaxExpandedCells` 是完整来源展开面积的保守上界，包含轴组合、小计、度量、总计和筛选占用，不按当前筛选缩小，也不等于进程内存限制。默认只检查格式行列边界。目标坐标是包括筛选区的整体左上角；OOXML 的透视主体 `location` 会位于筛选区下方。

**高级模式尚未完成最终验收。** 本机 Excel 16.0 / Build 20430 的真实夹具中，文本 Count 与 Minimum/Maximum、多列轴组合时，部分小计及总计为空；显式刷新后仍存在。全新 Excel 工作簿的原生创建对照也能复现，数值 Count 及部分简单组合正确，尚未定位可满足完整合同的修复。源码配置和包结构通过不能替代数值成功；不会改写为静态汇总或将文本 Count 替换为数值 Count 来宣布完成。详细证据保存在 [高级透视执行记录](../../ai_docs/tasks/BING-OFFICES-ADVANCED-PIVOT-TABLES-20261005-001/execution.md)。

NPOI XLSX 复用相同公共定义，在全部业务 Sheet 写入后预检并创建透视；透视请求使用 XSSF DOM，避免 SXSSF 已刷新行无法回读。缓存记录与内联校验修正共用一次磁盘暂存，保留全部来源及原始类型，不预写静态聚合值。显式保留模式支持已有透视模板及业务 Sheet 桥接；XLS/HSSF 仍拒绝透视，缓存刷新仍拒绝。NPOI 文件结构与公开重开成功不代表 Excel 首开、刷新或其他客户端已通过；当前验证进度见 [NPOI 执行记录](../../ai_docs/tasks/BING-OFFICES-NPOI-PIVOT-CACHE-20261005-001/execution.md)。

下面的示例展示完整高级请求，并用 ClosedXML 同步、异步导出后重开检查业务来源、透视字段和缓存记录。它验证文件结构，不验证客户端计算出的聚合数值；特别是文本 Count 的 Excel 数值行为仍未通过最终验收。

```csharp
var rows = new[]
{
    new WarehousePivotRow { Warehouse = "North", Category = "Food", Month = "Jan", Status = "Open", Amount = 10m, Item = "A" },
    new WarehousePivotRow { Warehouse = "North", Category = "Food", Month = "Feb", Status = "Closed", Amount = 30m, Item = "B" },
    new WarehousePivotRow { Warehouse = "North", Category = "Tools", Month = "Jan", Status = "Open", Amount = 20m, Item = "C" },
    new WarehousePivotRow { Warehouse = "South", Category = "Food", Month = "Jan", Status = "Pending", Amount = 40m, Item = "D" }
};
var request = ExcelExport.Workbook(workbook => workbook
    .AddSheet("Data", rows)
    .PivotTable(new ExcelPivotTableDefinition
    {
        Name = "WarehouseReport",
        SourceSheet = "Data",
        SourceRange = new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0,
            EndRow = rows.Length, EndColumn = 5
        },
        TargetSheet = "Summary",
        TargetRow = 1, TargetColumn = 2,
        RowFields = new[]
        {
            new ExcelPivotAxisFieldDefinition { Field = "Warehouse", SortOrder = ExcelPivotSortOrder.Ascending },
            new ExcelPivotAxisFieldDefinition { Field = "Category", SortOrder = ExcelPivotSortOrder.Descending,
                SubtotalMode = ExcelPivotSubtotalMode.None }
        },
        ColumnFields = new[]
        {
            new ExcelPivotAxisFieldDefinition { Field = "Month", SubtotalMode = ExcelPivotSubtotalMode.None }
        },
        ValueFields = new[]
        {
            new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total", NumberFormat = "0.00" },
            new ExcelPivotValueDefinition { Field = "Amount", Caption = "Mean", Aggregation = ExcelPivotAggregation.Average,
                NumberFormat = "0.0" },
            new ExcelPivotValueDefinition { Field = "Amount", Caption = "Low", Aggregation = ExcelPivotAggregation.Minimum,
                NumberFormat = "0.000\"kg\"" },
            new ExcelPivotValueDefinition { Field = "Amount", Caption = "High", Aggregation = ExcelPivotAggregation.Maximum,
                NumberFormat = "0.00" },
            new ExcelPivotValueDefinition { Field = "Item", Caption = "Count", Aggregation = ExcelPivotAggregation.Count }
        },
        Filters = new[]
        {
            new ExcelPivotFilterDefinition { Field = "Status", Values = new object[] { "Open", "Pending" } }
        }
    }));
var exporter = new ClosedXmlExcelExporter();
using var syncOutput = new MemoryStream();
exporter.Export(request, syncOutput);
using var asyncOutput = new MemoryStream();
await exporter.ExportAsync(request, asyncOutput);

void VerifyWorkbook(byte[] bytes)
{
    using var workbookStream = new MemoryStream(bytes);
    using var workbook = new XLWorkbook(workbookStream);
    var source = workbook.Worksheet("Data").Range("A1:F5").Cells()
        .Select(cell => cell.GetString()).ToArray();
    var expectedSource = new[]
    {
        "Warehouse", "Category", "Month", "Status", "Amount", "Item",
        "North", "Food", "Jan", "Open", "10", "A",
        "North", "Food", "Feb", "Closed", "30", "B",
        "North", "Tools", "Jan", "Open", "20", "C",
        "South", "Food", "Jan", "Pending", "40", "D"
    };
    if (!source.SequenceEqual(expectedSource))
        throw new InvalidOperationException("导出工作簿未保留完整来源数据。");

    var pivot = workbook.Worksheet("Summary").PivotTables.Single();
    if (pivot.Name != "WarehouseReport"
        || !pivot.RowLabels.Select(field => field.SourceName).SequenceEqual(new[] { "Warehouse", "Category" })
        || pivot.RowLabels.Get("Warehouse").SortType != XLPivotSortType.Ascending
        || pivot.RowLabels.Get("Category").SortType != XLPivotSortType.Descending
        || !pivot.RowLabels.Get("Warehouse").Subtotals.Contains(XLSubtotalFunction.Automatic)
        || pivot.RowLabels.Get("Category").Subtotals.Any()
        || !pivot.ColumnLabels.Any(field => field.SourceName == "Month")
        || !pivot.Values.Select(value => value.CustomName)
            .SequenceEqual(new[] { "Total", "Mean", "Low", "High", "Count" })
        || !pivot.Values.Select(value => value.SummaryFormula).SequenceEqual(new[]
        {
            XLPivotSummary.Sum, XLPivotSummary.Average, XLPivotSummary.Minimum,
            XLPivotSummary.Maximum, XLPivotSummary.Count
        })
        || pivot.ReportFilters.Single().SourceName != "Status"
        || !pivot.ReportFilters.Single().SelectedValues.Select(value => value.ToString())
            .SequenceEqual(new[] { "Open", "Pending" }))
        throw new InvalidOperationException("导出工作簿未保留高级透视配置。");

    using var package = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
    var cacheDefinitionEntry = package.Entries.Single(entry =>
        entry.FullName.Split('/').Last().StartsWith("pivotCacheDefinition", StringComparison.Ordinal)
        && entry.FullName.EndsWith(".xml", StringComparison.Ordinal));
    using var definitionStream = cacheDefinitionEntry.Open();
    var cacheDefinition = System.Xml.Linq.XDocument.Load(definitionStream);
    XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    var fields = cacheDefinition.Root.Element(spreadsheet + "cacheFields").Elements().ToArray();
    if (!fields.Select(field => (string)field.Attribute("name")).SequenceEqual(new[]
        { "Warehouse", "Category", "Month", "Status", "Amount", "Item" }))
        throw new InvalidOperationException("透视缓存未保留全部来源字段。");

    var cacheRecordsEntry = package.Entries.Single(entry =>
        entry.FullName.Split('/').Last().StartsWith("pivotCacheRecords", StringComparison.Ordinal)
        && entry.FullName.EndsWith(".xml", StringComparison.Ordinal));
    using var recordsStream = cacheRecordsEntry.Open();
    var recordsDocument = System.Xml.Linq.XDocument.Load(recordsStream);
    var records = recordsDocument.Root.Elements(spreadsheet + "r").ToArray();
    var cachedSource = records.SelectMany(record => record.Elements().Select((cell, index) =>
    {
        var value = (string)cell.Attribute("v");
        return cell.Name == spreadsheet + "x"
            ? (string)fields[index].Element(spreadsheet + "sharedItems").Elements()
                .ElementAt(int.Parse(value)).Attribute("v")
            : value;
    })).ToArray();
    if (records.Length != rows.Length || records.Any(record => record.Elements().Count() != fields.Length)
        || !cachedSource.SequenceEqual(expectedSource.Skip(fields.Length)))
        throw new InvalidOperationException("透视缓存未保留所有来源记录，包括被筛选的记录。");

    var pivotEntry = package.Entries.Single(entry =>
        entry.FullName.Split('/').Last().StartsWith("pivotTable", StringComparison.Ordinal)
        && entry.FullName.EndsWith(".xml", StringComparison.Ordinal));
    using var pivotStream = pivotEntry.Open();
    var pivotDocument = System.Xml.Linq.XDocument.Load(pivotStream);
    var dataFields = pivotDocument.Root.Element(spreadsheet + "dataFields").Elements().ToArray();
    if (!dataFields.Select(field => (string)field.Attribute("name"))
        .SequenceEqual(new[] { "Total", "Mean", "Low", "High", "Count" }))
        throw new InvalidOperationException("透视缓存未保留所有度量标题。");

    using var stylesEntry = package.GetEntry("xl/styles.xml").Open();
    var styles = System.Xml.Linq.XDocument.Load(stylesEntry);
    var numberFormats = styles.Root.Element(spreadsheet + "numFmts")?.Elements()
        .ToDictionary(item => (string)item.Attribute("numFmtId"), item => (string)item.Attribute("formatCode"))
        ?? new Dictionary<string, string>();
    var actualFormats = dataFields.Take(4).Select(field =>
    {
        var id = (string)field.Attribute("numFmtId");
        if (id == "2")
            return "0.00";
        return numberFormats.TryGetValue(id, out var format) ? format : null;
    });
    if (!actualFormats.SequenceEqual(new[] { "0.00", "0.0", "0.000\"kg\"", "0.00" }))
        throw new InvalidOperationException("透视度量未保留指定数字格式。");
}

VerifyWorkbook(syncOutput.ToArray());
VerifyWorkbook(asyncOutput.ToArray());

/// <summary>
/// 表示高级透视示例中的仓库业务记录。
/// </summary>
public sealed class WarehousePivotRow
{
    /// <summary>
    /// 初始化一个 <see cref="WarehousePivotRow"/> 类型的实例。
    /// </summary>
    public WarehousePivotRow()
    {
    }

    /// <summary>
    /// 获取或设置仓库名称。
    /// </summary>
    public string Warehouse { get; set; }

    /// <summary>
    /// 获取或设置商品类别。
    /// </summary>
    public string Category { get; set; }

    /// <summary>
    /// 获取或设置月份。
    /// </summary>
    public string Month { get; set; }

    /// <summary>
    /// 获取或设置采购状态。
    /// </summary>
    public string Status { get; set; }

    /// <summary>
    /// 获取或设置金额。
    /// </summary>
    public decimal Amount { get; set; }

    /// <summary>
    /// 获取或设置商品标识。
    /// </summary>
    public string Item { get; set; }
}
```

基础透视的下述客户端成功记录仅覆盖旧单字段 Sum；不推广为高级模式成功，也不代表已有透视模板可更新或刷新。模板保留仅适用于显式 `PreserveTemplatePivotTables()` 模式。

## 客户端显示与缓存生命周期

Provider 文件包含透视定义与真实缓存，没有预先写入普通汇总单元格。依赖普通单元格值的无客户端预览不能因此得到静态汇总；这与 Excel 中的透视显示是不同合同。

本机 Excel 16.0、Build 20430 的历史客户端验证仅覆盖基础透视新建的 net6.0/net8.0 同步及异步夹具：完整来源为 North 10、South 20、North 30，首次透视显示 North 40、South 20、总计 60。将首笔金额改为 15 后，显式刷新前仍显示旧结果；刷新 PivotCache 和 PivotTable 后显示 North 45、South 20、总计 65，保存验证副本并重开后结果保持正确。缓存来源及字段角色也参与断言。该证据不覆盖本任务的模板更新、更新后来源变化或新建与更新组合。

本次观察到 `RefreshOnFileOpen = true`，但不能仅凭该标志位推断 COM 打开或其它客户端自动刷新。Microsoft 对 [RefreshOnFileOpen](https://learn.microsoft.com/office/vba/api/excel.pivotcache.refreshonfileopen) 的说明区分了自动打开和对象模型的 Open；需要确定来源更新后的结果时，应显式刷新并验证完整结果。这里没有新增公共刷新 API；Provider 的显式保留模式只回挂已有透视部件，更新子集会为新缓存写入打开时刷新请求，但两种路径都不执行客户端刷新。

模板保留/更新结构、缓存刷新标记以及业务组合合同由文件级测试验证；本轮不运行 Excel、WPS 或 LibreOffice 客户端。客户端刷新、首次显示及保存重开行为不据包结构或历史基础透视夹具外推。

客户端探针保留正常模式、原始矩阵、刷新前后状态及清理结果。受限会话曾出现 COM 登录会话错误；另一次初始读取返回空来源及空矩阵，重试取得完整结果。两份 Failed 记录均保留，尚未证明空快照根因；维护脚本可被动观察初始化，超时或错误结果仍失败，不主动刷新来冒充首次显示。

模板更新负责写入新的来源缓存和透视关系，并设置新缓存的打开时刷新请求；客户端刷新、首开显示和保存重开仍需单独验证，不能由包结构通过推断。

| 验证对象 | 当前证据 |
| --- | --- |
| 完整来源、定义和缓存 | 双 TFM 同步/异步夹具及职责测试 |
| Excel 首次显示、显式刷新、保存重开 | 上述本机版本的真实 COM 结果；不推广为所有版本保证 |
| WPS / LibreOffice | 未验证，当前未发现本机客户端环境 |
| 无客户端静态汇总 | 未提供普通汇总单元格结果 |

生成和验收步骤见 [维护探针](../../build/PivotClientProbe/README.md)，详细范围及失败记录见 [执行记录](../../ai_docs/tasks/BING-OFFICES-PIVOT-CLIENT-LIFECYCLE-20261005-001/execution.md)。每次验收使用本任务生成的新副本，不访问用户活动工作簿。

## 多透视资源边界

本机 1K/10K 行、1/4 个透视的有界矩阵中，相同来源的四个透视共享一个缓存，业务序列仅枚举一次；四个不同来源则分别建立缓存。业务枚举计数不等于 DOM 预检扫描次数，不能据此宣称整个流程只扫描一次。

来源枚举阶段的双 TFM 同步/异步取消均保留既有文件并清理暂存；这些响应时间不代表缓存构建和序列化阶段。资源数据、测量区间和后续建议见 [资源报告](../../ai_docs/tasks/BING-OFFICES-PIVOT-CLIENT-LIFECYCLE-20261005-001/resource-report.md)。更多字段、透视和来源会增加 DOM、分配及缓存开销，应结合部署资源单独评估。

## 核心实现与依赖配置

公共定义位于 `src/Bing.Offices.Abstractions/Bing/Offices/Exports/Reports/ExcelPivot*.cs`；NPOI 的 `Exports/PivotTables` 负责来源验证、目标预算、定义与完整缓存写入，ClosedXML 的对应目录及 `ClosedXmlExcelExporter.cs` 负责引擎适配和缓存修正。模板桥接/清单/缓存部件位于 `src/provider-shared/ExcelPivotTemplate*.cs`。依赖匹配的 NPOI 或 ClosedXML 包；没有独立客户端刷新服务。高级聚合及模板更新客户端结果统一标为**待实现/待确认（待确认）**，不把历史基础验收外推为当前所有组合通过。
