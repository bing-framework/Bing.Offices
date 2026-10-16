# NuGet 迁移

当前工作树包版本为 `2.0.0`，核心三包的身份和依赖方向保持不变（其余 Provider、分析器与合同包见模块总览）：

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

旧的 `OfficeException` 异常层级、`ExcelSetting`/`SheetSetting` 已从发布 API 移除；CSV 实体实现类不再作为发布 API，调用方应通过 `ICsvImporter`/`ICsvExporter` 或 DI 使用。本轮另经成员级批准移除四个重复的 `ExportToFile`/`ExportToFileAsync` 扩展声明，接口实例成员保持不变。历史 `dotnet pack` 和本地包消费者记录只能证明当时被测候选的可消费性，不证明当前工作区已重新通过，也不证明公共 NuGet 发布完成。

## API 迁移对照

下表区分历史入口、当前真实替代项和兼容状态；审批列属于历史任务记录，不作为本轮新增审批。未实现的 API 名称不列为可调用替代项：

| 历史或当前入口 | 当前可用替代项 | 当前兼容策略 | 历史记录 |
| --- | --- | --- | --- |
| `ExcelMapping.For<T>()` | `ImportMappingBuilder<T>` / `ExportMappingBuilder<T>` | 已移除 | 本次 RC Breaking Change |
| `Mapping(configuration)` / `Mapping(document)` | 继续使用实际存在的 `Mapping(...)` 重载 | 当前公开入口 | `MappingConfiguration(...)` / `MappingDocument(...)` 具名方法为待实现/待确认，不能调用 |
| `HeaderMatch` | `RequireExpectedHeaders` | 删除 | Major 已迁移 |
| `MaxColumnCount` | `MaxReadColumns` | 删除 | Major 已迁移 |
| `EnabledEmptyLine` | `ReportEmptyRows` | 删除 | Major 已迁移 |
| `IgnoreEmptyLineAfterData` | `StopAtFirstEmptyRow` | 删除 | Major 已迁移 |
| `AddNavigationSheet` | `AddSheet(name, parents.SelectMany(...))` | 保留 | 待批准 |
| `ExcelSetting` / `SheetSetting` | `ExcelWorkbookMetadataOptions` 与 Sheet request | 已移除 | Round 6 已批准 |
| `OfficeException` 异常层级 | `BingOfficesException` 家族、标准参数/取消异常及结构化导入错误 | 已移除 | Round 6 已批准 |
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

需要按打印页输出中间小计时，在同一列表区域配置 `PageBreak(rowsPerPage)` 和 `PageSubtotal`。分页小计只对存在下一页的完整明细页生效，最后一页继续使用普通 `Footer`；导入会识别 marker 并跳过小计区域。可同时配置 `GroupSubtotal`：页容量只计算明细，同一边界先写分组小计再写分页小计。NPOI 的 XLS/XLSX 和 ClosedXML 的 XLSX 均支持。

旧单据若只在打印时重复已有表头，可在 Entity 布局上调用 `PrintTitleRows("明细", "1:3")`，无需移动明细行。旧代码若逐页复制单据头，则在同一列表区域配置 `PageBreak` 和 `PageHeader("A1:D3", "A1", "页眉", header => header.Cell("D1", page => $"第 {page.PageNumber}/{page.PageCount} 页"))`；示例要求列表表头从 `B3` 开始，首 页源区域末行紧邻首个明细行。`PageHeader` 会在后续逻辑页插入可供导入跳过的实体行，页容量仍只统计明细。同一 Sheet 的两种模式互斥，包括模板已有的打印标题；每个 Sheet 至多一个实际重复页眉区域，可同时配置不重叠的普通列表。迁移模板时需确认页眉的 marker、完整表头、合并区域及预留目标行没有冲突；公式只保留文本，不复制图片和图表。详见 [Entity 分页与重复页眉](dynamic-columns.md#重复打印标题与实际页眉)。

旧 NPOI 代码若仅为按地区汇总金额而手动创建透视表，可在 Workbook 构建时增加 `PivotTable(new ExcelPivotTableDefinition { ... })`，将来源 Sheet、含表头的有界区域、行字段、数值字段与目标 Sheet 显式声明。ClosedXML XLSX 与 NPOI XLSX 支持新建基础及高级透视；高级透视的客户端数值 Count 尚未完成验收。NPOI XLS/HSSF 不支持该公共透视请求，迁移前先核对 Provider 与文件格式。示例和边界见 [Workbook 透视表](pivot-tables.md)。
### Entity 独立逐页锚点迁移

旧代码需要在每个逻辑页保存独立页眉或签字区时，先配置 `PageBreak`，再选择 `PageHeaderNamed`、`PageFooter` 或 `PageFooterNamed`。下面沿用本页已有的 `PurchaseOrder`/`PurchaseLine` DTO；它是可编译的布局片段，名称策略显式使用默认安全值：

```csharp
var layout = ExcelEntity.Layout<PurchaseOrder>(builder => builder
    .ListRegion("采购单", "A5", order => order.Lines, region => region
        .PageBreak(20)
        .PageHeaderNamed("Head", "A1:H5", "A1", "页眉")
        .PageFooterNamed("Sign", "签字", footer => footer
            .GapRows(1)
            .Cell("B1", lines => lines.Count)
            .Cell("A2", "审核"),
            ExcelEntityPageAnchorStaleNamePolicy.Reject)
        .Footer("合计", footer => footer
            .Cell("C1", lines => lines.Sum(line => line.Quantity * line.UnitPrice),
                numberFormat: "0.00"))));

var order = new PurchaseOrder
{
    Lines = new List<PurchaseLine>
    {
        new PurchaseLine { Name = "键盘", Quantity = 2, UnitPrice = 12.5m },
        new PurchaseLine { Name = "鼠标", Quantity = 3, UnitPrice = 7m }
    }
};
foreach (var useXls in new[] { true, false })
{
    using IWorkbook templateWorkbook = useXls ? new HSSFWorkbook() : new XSSFWorkbook();
    templateWorkbook.CreateSheet("采购单");
    using var template = new MemoryStream();
    templateWorkbook.Write(template, true);
    template.Position = 0;
    using var output = new MemoryStream();
    new NpoiExcelExporter().ExportForTemplate(order, layout,
        new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

    output.Position = 0;
    var imported = new NpoiExcelImporter().ImportEntity(output, layout);
    if (!imported.IsSuccess || !order.Lines.Select(line => (line.Name, line.Quantity, line.UnitPrice))
        .SequenceEqual(imported.Entity.Lines.Select(line => (line.Name, line.Quantity, line.UnitPrice))))
        throw new InvalidOperationException("逐页锚点示例读回失败。");
    using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false));
    var sheet = workbook.GetSheet("采购单");
    if (sheet.LastRowNum != 10 || sheet.GetRow(8).GetCell(0).StringCellValue != "签字"
        || sheet.GetRow(8).GetCell(1).NumericCellValue != 2
        || sheet.GetRow(9).GetCell(0).StringCellValue != "审核"
        || sheet.GetRow(10).GetCell(2).NumericCellValue != 46
        || !new[] { "Head_001", "Sign_001" }.OrderBy(name => name)
            .SequenceEqual(workbook.GetAllNames().Select(name => name.NameName).OrderBy(name => name)))
        throw new InvalidOperationException("逐页签字区结构不符合预期。");
}
```

此例默认在 `A5` 写入列表表头、从第 6 行写明细，因此页眉源范围必须结束于第 5 行；`A1:H5` 同时预留了单据头和完整表头宽度。最终金额按明细的数量乘单价计算，签字区不参与聚合。需要原生公式时，继续使用已有 `FormulaSumDetailRowsAbove` 并明确实际明细金额列。

`PageHeaderNamed` 与 `PageFooterNamed` 按逻辑页生成首个、末个以及空/单页名称；每页 Footer 在最终 Footer 之前写入，marker 列仍是导入扫描边界。模板清理策略 `RemoveObsolete` 只允许删除当前系列中合法且未使用的旧单格名称；跨 Sheet、二维、公式、失效或同名多作用域引用仍拒绝并保护输入模板及旧输出。NPOI 的 XLS/HSSF、XLSX/XSSF 与 ClosedXML 的 XLSX 有本阶段职责测试；这里不把 WPS、LibreOffice、MiniExcel、SpreadCheetah 或 ExcelDataReader 的验证状态扩大为支持，也不宣称未运行的门禁已通过。

透视来源区域只接受字面值。旧代码若依赖读取来源单元格时隐式计算公式，应先由应用或显式计算接口求值，再将结果写入来源；本入口会拒绝来源内任何公式及 Excel 错误，包括带新鲜、过期或缺失缓存的模板公式。求和字段须为真实数值，数字文本、日期和布尔值不能隐式参与求和；来源错误通过 `ConfigurationInvalid/Preflight` 返回原始一基行列。公开签名保持兼容，此处收紧输入行为，升级时需要检查来源数据。

迁移时还需区分完整透视缓存与普通汇总单元格。本入口没有写出无客户端静态汇总；本机 Excel 的历史客户端夹具仅验证了基础透视新建后的首次显示、显式刷新和保存重开，不覆盖高级透视；高级透视（尤其文本 Count）尚未完成客户端验收，也不据此承诺 WPS/LibreOffice 或所有 Excel 版本。来源值更新后不要只依赖刷新标志位，需由客户端显式刷新并核对结果。可运行的包消费者与维护步骤见 [透视探针](../../build/PivotClientProbe/README.md)。

模板中如需通过工作簿名称定位最终合计行，将原有 `Footer("合计", ...)` 改为 `FooterNamed("OrderTotal", "合计", ...)`。导出时，NPOI 的 XLS/XLSX 和 ClosedXML 的 XLSX 会把 `OrderTotal` 指向本次最终 marker 单元格；模板中同名、同 Sheet 的单格名称随明细长度更新，缺失时创建。需要定位每个非末页小计时，使用 `PageSubtotalNamed("PageTotal", "页小计", ...)`，生成 `PageTotal_001` 起的逐页名称。页数缩短后模板若保留未使用的同系列名称会报 Plan 错误，需同步清理模板。导入仍使用 marker 文本；名称属于工作簿或目标 Sheet，跨 Sheet、范围名称或重名在导出预检中返回 Plan 配置错误。

## 实现定位、依赖与升级限制

包 ID、TFM 和依赖由 `src/*/*.csproj`、`framework.props`、`version.props` 与 `version.dev.props` 确定。方向 Profile 和请求契约在 Abstractions，映射/流扩展在 Core，NPOI 用户扩展在 NPOI 的 `Bing/Offices/Extensions`。API 快照工具位于 `build/ApiSnapshot`，测试与包消费者入口见[构建与验证](../development/build-validation.md)。

升级应配套选择 Core/Abstractions/Provider，避免旧二进制继续依赖已移除的内部友元或类型。包源中实际可用版本与当前候选门禁为**待实现/待确认（待确认）**，不要把仓库版本号当作线上状态。
