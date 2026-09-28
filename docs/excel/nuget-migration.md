# NuGet Migration

当前工作树包版本为 `2.0.0`，三个包身份和依赖方向保持不变：

`Bing.Offices.Abstractions <- Bing.Offices.Core <- Bing.Offices.Npoi`

迁移建议：

1. 新代码使用四种方向明确的 Profile 契约，按方向配置 Builder。
2. JSON/XML 只使用 v2 `ExcelMappingDocument`；v1 平铺配置应在升级前离线转换。
3. 使用 `ExcelRequired`、`ExcelRegex`、`ExcelDate`、`ExcelMaxValue`、`ExcelRange`、`ExcelMaxLength`、`ExcelUnique`。
4. 通过 `AddBingOfficesNpoi(IServiceCollection): IServiceCollection` 注册 NPOI，并支持链式注册；Profile Registry 使用独立的显式或程序集扫描扩展。
5. 只依赖 provider-neutral 请求、结果和转换器接口，不引用 NPOI 类型。
6. 调用方提供的输入、输出和配置流仍由调用方拥有，库不会关闭它们。

当前 `2.0.0` 工作树已移除旧的 `ExcelSetting`/`SheetSetting` 配置类型；Workbook metadata 改为请求级快照。CSV 新增输入字节、行、错误、字段和列限制，超限结果使用 `CsvImportErrorCode.ResourceLimit` 并设置 `CsvImportResult.IsTruncated`。

## 旧版调用迁移提示

可选分析器包对仍引用 1.x 的项目提供两条 Warning：`BOM001` 识别旧版 `IExcelExportService` / `ExportOptions<T>`，`BOM002` 识别 `[HasDynamicColumn]`。先在旧项目中启用分析器、定位调用，再升级依赖；旧符号在升级后无法解析时，分析器不会仅凭名称猜测。

旧导出服务的 `ExportAsync(options)` 需要按报表的 Sheet、列、模板和流所有权语义重建为 `ExcelExport.Workbook(...)` 请求，再交给 `IExcelExporter`。旧模型含多个 `[DynamicColumn]` 字典时，应为每个字典在 Entity 列表区域配置独立的 `DynamicColumnGroup(...)`、稳定 Key 和标题。2.x 仍保留单字典 `[DynamicColumn]` 及 `ColumnNameAttribute`，不必仅因它们存在而迁移。

部分业务项目自行定义 `ImportLocationAttribute`，通过行列坐标和反射读取固定单元格。这不是 Bing.Offices 的通用旧 API；迁移时可用 `[ExcelEntityCell(sheetName, address)]` 与 `ExcelEntity.LayoutFromAttributes`，但必须人工核对 A1 坐标、工作表和原有校验，分析器不会自动转换。

失败工作簿输出限制已从 `MaxBytes` 更名为 `MaxSerializedBytes`。该限制只约束失败工作簿序列化输出，不代表原始 Workbook、解压内容、实体或失败输出 DOM 的内存上限。

旧的 `OfficeException` 异常层级、`ExcelSetting`/`SheetSetting` 已从发布 API 移除；CSV 实体实现类不再作为发布 API，调用方应通过 `ICsvImporter`/`ICsvExporter` 或 DI 使用。本轮另经成员级批准移除四个重复的 `ExportToFile`/`ExportToFileAsync` 扩展声明，接口实例成员保持不变。`dotnet pack` 和本地 package consumer 验证确认当前 `2.0.0` 包的可消费性。

## API 迁移对照

下表列出仍保持兼容的旧版入口和候选替代项；已标记移除的类型不应继续在新代码中引用：

| 当前 2.x 入口 | 新代码建议 | 当前兼容策略 | 批准状态 |
| --- | --- | --- | --- |
| `ExcelMapping.For<T>()` | `ImportMappingBuilder<T>` / `ExportMappingBuilder<T>` | 已移除 | 本次 RC Breaking Change |
| `Mapping(configuration)` / `Mapping(document)` | 使用具名的 `MappingConfiguration(...)` / `MappingDocument(...)` 入口 | 迁移中 | 当前 Major 仍使用 `Mapping(...)` |
| `HeaderMatch` | `RequireExpectedHeaders` | 删除 | Major 已迁移 |
| `MaxColumnCount` | `MaxReadColumns` | 删除 | Major 已迁移 |
| `EnabledEmptyLine` | `ReportEmptyRows` | 删除 | Major 已迁移 |
| `IgnoreEmptyLineAfterData` | `StopAtFirstEmptyRow` | 删除 | Major 已迁移 |
| `AddNavigationSheet` | `AddSheet(name, parents.SelectMany(...))` | 保留 | 待批准 |
| `ExcelSetting` / `SheetSetting` | `ExcelWorkbookMetadataOptions` 与 Sheet request | 已移除 | Round 6 已批准 |
| `OfficeException` 异常层级 | 标准参数/状态异常与结构化导入错误 | 已移除 | Round 6 已批准 |
| `CsvEntityImporter` / `CsvEntityExporter` | `ICsvImporter` / `ICsvExporter`，优先通过 DI 获取 | 已 internal 化 | Round 6 已批准 |
| `ICellValueConverter` | `IExcelValueConverter` | 已移除；请迁移到提供程序无关的双向转换器 | 本任务已批准 |
| `CsvStreamExtensions.ExportToFile*` / `ExcelStreamExtensions.ExportToFile*` | `ICsvExporter.ExportToFile*` / `IExcelExporter.ExportToFile*` | 扩展声明已移除；接口实例成员保留 | FIX-003，`approvedBy=jian玄冰` |

迁移示例：

```csharp
// 使用方向明确的 Builder/Profile 配置
var importMapping = new ImportMappingBuilder<OrderRow>()
    .Property(row => row.Code).HasHeader("订单号").And()
    .Build();
services.AddMappingProfile<OrderProfile>();
var request = ExcelImport.Workbook<OrderWorkbook>(builder =>
	builder.Sheet("订单", workbook => workbook.Rows));
```

请在每个 Workbook 请求上调用 `Metadata(...)`；不配置时使用请求级默认值。CSV 调用方应依赖 `ICsvImporter`/`ICsvExporter`，不要直接构造 Core 内部实现。对于已移除的四个文件导出扩展，直接改为接口实例调用：`exporter.ExportToFile(...)` 或 `await exporter.ExportToFileAsync(...)`。其余未列入本次批准范围的入口保持兼容。

## NPOI 单据代码迁移到 Entity Layout

旧式 NPOI 单据通常在策略中创建工作簿、按行写入固定列、手动合并单元格并在明细循环后写合计。迁移时保留业务模型，把固定单元格和明细区域声明为布局；Provider 负责创建 Sheet、转换值、校验和写入目标流。

下面的示例可放入引用 `Bing.Offices.Npoi` 的控制台项目运行。它需要 `using System;`、`using System.Collections.Generic;`、`using System.IO;`、`using System.Linq;`、`using Bing.Offices.Entities;`、`using Bing.Offices.Exports;`、`using Bing.Offices.Imports;`、`using NPOI.HSSF.UserModel;`、`using NPOI.SS.UserModel;` 和 `using NPOI.XSSF.UserModel;`。`ExportEntity` 创建 XLSX；要生成 XLS，则以 HSSF 模板作为格式来源。示例还分别填充 HSSF/XLS 与 XSSF/XLSX 模板、重新导入，并检查明细和合计。

```csharp
public static class PurchaseOrderMigrationExample
{
    public static void Main()
    {
        var order = new PurchaseOrder
        {
            Number = "PO-2026-001",
            Customer = "华东采购部",
            Lines = new List<PurchaseLine>
            {
                new PurchaseLine { Name = "键盘", Quantity = 2, UnitPrice = 12.5m },
                new PurchaseLine { Name = "鼠标", Quantity = 3, UnitPrice = 7m }
            }
        };

        var layout = ExcelEntity.LayoutFromAttributes<PurchaseOrder>(builder => builder
            .ListRegion("采购单", "A5", item => item.Lines, region => region
                .Footer("合计", footer => footer
                    .Cell("C1", lines => lines.Sum(line => line.Quantity * line.UnitPrice),
                        numberFormat: "0.00"))));
        var exporter = new NpoiExcelExporter();
        var importer = new NpoiExcelImporter();

        // 非模板实体导出新建 XLSX 工作簿。
        using var xlsx = new MemoryStream();
        exporter.ExportEntity(order, layout, xlsx);
        AssertImportedOrder(importer.ImportEntity(
            new MemoryStream(xlsx.ToArray(), writable: false), layout), order);
        AssertFooter(xlsx.ToArray(), expectedTotal: 46m);

        // 模板决定容器格式；同一布局可以填充 HSSF/XLS 与 XSSF/XLSX 模板。
        foreach (var useXls in new[] { true, false })
        {
            var templateBytes = CreateTemplate(useXls);
            using var template = new MemoryStream(templateBytes, writable: false);
            using var output = new MemoryStream();
            exporter.ExportForTemplate(order, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

            using (var formatStream = new MemoryStream(output.ToArray(), writable: false))
            using (var workbook = WorkbookFactory.Create(formatStream))
            {
                if ((useXls && !(workbook is HSSFWorkbook))
                    || (!useXls && !(workbook is XSSFWorkbook)))
                    throw new InvalidOperationException("导出格式与模板格式不一致。");
            }

            using var importSource = new MemoryStream(output.ToArray(), writable: false);
            using var importTemplate = new MemoryStream(templateBytes, writable: false);
            var result = importer.ImportForTemplate(importSource, layout,
                new ExcelEntityTemplateOptions(importTemplate, leaveOpen: true));
            AssertImportedOrder(result, order);
            AssertFooter(output.ToArray(), expectedTotal: 46m);
        }
    }

    private static byte[] CreateTemplate(bool useXls)
    {
        using var workbook = useXls
            ? (IWorkbook)new HSSFWorkbook()
            : new XSSFWorkbook();
        workbook.CreateSheet("采购单");
        using var stream = new MemoryStream();
        workbook.Write(stream, true);
        return stream.ToArray();
    }

    private static void AssertImportedOrder(ExcelEntityImportResult<PurchaseOrder> result,
        PurchaseOrder expected)
    {
        if (!result.IsSuccess)
            throw new InvalidOperationException(string.Join("; ", result.Errors.Select(error => error.Message)));
        if (result.Entity.Number != expected.Number || result.Entity.Customer != expected.Customer
            || result.Entity.Lines.Count != expected.Lines.Count
            || !result.Entity.Lines.Zip(expected.Lines, (actual, source) =>
                actual.Name == source.Name && actual.Quantity == source.Quantity
                    && actual.UnitPrice == source.UnitPrice).All(matches => matches))
            throw new InvalidOperationException("导入的采购单或明细与导出数据不一致。");
    }

    private static void AssertFooter(byte[] bytes, decimal expectedTotal)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var workbook = WorkbookFactory.Create(stream);
        var sheet = workbook.GetSheet("采购单");
        var footer = sheet.GetRow(7);
        if (footer?.GetCell(0)?.StringCellValue != "合计"
            || Convert.ToDecimal(footer.GetCell(2).NumericCellValue) != expectedTotal)
            throw new InvalidOperationException("采购单合计行内容错误。");
    }
}

public sealed class PurchaseOrder
{
    [ExcelEntityCell("采购单", "B2")]
    public string Number { get; set; }

    [ExcelEntityCell("采购单", "B3")]
    public string Customer { get; set; }

    public List<PurchaseLine> Lines { get; set; } = new List<PurchaseLine>();
}

public sealed class PurchaseLine
{
    public string Name { get; set; }

    public int Quantity { get; set; }

    public decimal UnitPrice { get; set; }
}
```

商品、产品和扩展字段分别位于多个字典时，在同一个 `ListRegion` 上调用多个 `DynamicColumnGroup`；导入会将定义的列写回各自字典。自定义 Provider 只需实现公开的 Entity Layout SPI，不需要访问生产程序集内部成员。

明细已经按仓库或业务分类连续排序时，可以追加 `GroupSubtotal` 写入分组小计；它复用 Footer 的聚合单元格和样式，导入会跳过小计行。该能力依赖相邻分组顺序，不会替业务层重新排序。

需要按打印页输出中间小计时，在同一列表区域配置 `PageBreak(rowsPerPage)` 和 `PageSubtotal`。分页小计只对存在下一页的完整明细页生效，最后一页继续使用普通 `Footer`；导入会识别 marker 并跳过小计区域。分页小计与 `GroupSubtotal` 互斥，NPOI 的 XLS/XLSX 和 ClosedXML 的 XLSX 均支持。

模板中如需通过工作簿名称定位最终合计行，将原有 `Footer("合计", ...)` 改为 `FooterNamed("OrderTotal", "合计", ...)`。导出时，NPOI 的 XLS/XLSX 和 ClosedXML 的 XLSX 会把 `OrderTotal` 指向本次最终 marker 单元格；模板中同名、同 Sheet 的单格名称随明细长度更新，缺失时创建。导入仍使用 marker 文本，分页小计与分组小计不使用该名称。名称属于工作簿或目标 Sheet；跨 Sheet、范围名称或重名在导出预检中返回 Plan 配置错误。
