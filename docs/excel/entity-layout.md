# 实体布局：固定字段与多明细区域

## 功能说明

Entity Layout 将一个聚合对象的固定字段、多个明细集合及父子关系放到指定工作表位置，适合采购单、签字模板和分区报表。它与按 Sheet 输出整个集合的 Workbook 请求是两个入口，共用映射、转换、校验和错误模型。

NPOI 支持 XLS/HSSF 与 XLSX/XSSF 模板；无模板新建实体文件为 XLSX。ClosedXML 支持 XLSX。MiniExcel、ExcelDataReader、SpreadCheetah 不提供 Entity Layout；Starter 的该功能为**待实现/待确认（待实现）**。

## 用法示例

下例在已注册 NPOI 的服务容器 `services` 中解析常规接口，通过 Core 的实体扩展导出到 XLSX 流并重新导入；所需命名空间为 `System.Collections.Generic`、`System.IO`、`Bing.Offices.Entities`、`Bing.Offices.Exports`、`Bing.Offices.Imports`、`Bing.Offices.Extensions`、`Microsoft.Extensions.DependencyInjection`。

```csharp
var layout = ExcelEntity.Layout<PurchaseOrder>(book => book
    .Cell("采购单", "B2", order => order.Number)
    .ListRegion("采购单", "A5", order => order.Lines, region => region.Header()));
var order = new PurchaseOrder
{
    Number = "PO-001",
    Lines = new List<PurchaseLine> { new() { Code = "P-01", Quantity = 2 } }
};
var exporter = services.GetRequiredService<IExcelExporter>();
var importer = services.GetRequiredService<IExcelImporter>();
using var output = new MemoryStream();
exporter.ExportEntity(order, layout, output);
output.Position = 0;
var result = importer.ImportEntity(output, layout);

public sealed class PurchaseOrder
{
    public string Number { get; set; }
    public List<PurchaseLine> Lines { get; set; } = new();
}

public sealed class PurchaseLine
{
    public string Code { get; set; }
    public int Quantity { get; set; }
}
```

属性式布局使用 `[ExcelEntityCell("采购单", "B2")]` 和 `ExcelEntity.LayoutFromAttributes<PurchaseOrder>(...)`；特性只提供固定字段，列表和关系仍通过 Builder 声明。命名位置用 `CellNamed`、`ListRegionNamed`，模板名称必须解析为指定 Sheet 的单个单元格，局部名称优先于全局名称。

模板入口为 `ExportForTemplate(entity, layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), destination)`；导入有对应 `ImportForTemplate`。实体文件输出和字节扩展位于 Core 的 `ExcelEntityExtensions`，同样使用文件提交契约。

## 核心逻辑与文件位置

- `src/Bing.Offices.Abstractions/Bing/Offices/Entities/{Layout,Lists,Anchors}`：构建布局快照，检查地址、占位、列表、Footer 和名称配置。
- `src/Bing.Offices.Core/Bing/Offices/Extensions/ExcelEntityExtensions.cs`：对外扩展入口；实体接口在 Abstractions 的 `Entities/{Import,Export}`。
- `src/Bing.Offices.Npoi/Bing/Offices/Entities`：模板锚点解析、布局预检、固定字段/列表读写、行规划及资源约束。
- `src/Bing.Offices.ClosedXml/Bing/Offices/Entities`：XLSX 布局执行与批量导入。
- `src/provider-shared/ExcelEntityRowPlan.cs`、`ExcelEntitySubtotalImportScanner.cs`、`ExcelEntityPageAnchorNames.cs`：共享物理行计划、小计识别和逐页名称规则。

Provider 先解析模板位置和最终映射列，再执行固定单元格及列表区域。导入区分明细、复制页眉、小计与尾部 marker，不把签字数值读成明细。关系由显式 `HasMany` 配置绑定，不能依靠列名自动猜测导航属性。

## 依赖与配置

注册 NPOI 或 ClosedXML 后，从常规 `IExcelExporter`/`IExcelImporter` 调用 Core 的实体扩展。NPOI 没有直接注册两个完整实体 SPI 接口，不能使用 `GetRequiredService<IExcelEntityExporter>()` 获取；ClosedXML 另提供了对应 DI 别名。批次接口则由两者直接注册。`ExcelEntityImportOptions` 控制资源、行内失败策略、校验来源和可选 `FailureOptions`。命名转换器及动态列校验通过 Core Mapping Plan 解析。

动态列组、计算列、连续分组小计、分页小计、实际复制页眉、每页签字区及最终 Footer 的完整示例见[动态列与分页](dynamic-columns.md)。批次、SummaryOnly 与 Forward 边界见[批量导入](batch-import.md)，编译期常量检查见[分析器](entity-layout-analyzers.md)。

## 已知限制与注意事项

- 构建时只能验证已知地址；命名锚点、模板占位和实际数据长度仍需 Provider 预检。布局不可重叠，命名列表不能再配置绝对 `End`。
- `PrintTitleRows` 只设置打印元数据，`PageHeader` 才复制单元格；两者不能在同一 Sheet 共用。逻辑分页不等于打印机产生的物理页。
- 计算列仅导出时执行；Footer 原生公式不由库计算。组小计按连续组处理，不自动重新排序数据。
- 同步与异步模板操作仍存在 DOM 同步阶段；调用方流和模板的 `LeaveOpen` 需分别管理。
- Entity 接口被标记为隐藏的 Provider SPI，但已在包中公开；不能从编辑器默认不显示推断其未实现。
