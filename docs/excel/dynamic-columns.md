# Dynamic Columns

动态列使用稳定 `Key` 和 `Alias`，实体属性通常是 `IDictionary<string, object>`。固定列由 `PropertyName` 绑定，CSV 和 Excel 都支持固定列与动态列并存。

```csharp
var request = ExcelImport.Workbook<OrdersWorkbook>(builder =>
    builder.Sheet("订单", workbook => workbook.Items, sheet => sheet
        .DynamicColumns(row => row.Values, new[]
        {
            new ExcelDynamicColumnDefinition { Key = "region", Title = "区域" },
            new ExcelDynamicColumnDefinition { Key = "channel", Title = "渠道" }
        })
        .Mapping(document)));
```

导出动态列使用相同的稳定 Key。未知动态值可选择失败或忽略；列标题、Key 和 Alias 必须在绑定阶段校验冲突。动态列的校验、空白策略和 Unique 规则沿用统一 Mapping 配置。

## Entity Layout 动态分组

实体列表区域可以把多个直接字典属性声明为独立动态列组。各组的 Key、标题和别名在该列表区域内统一校验，导入时按定义回写到所属字典；未配置分组时，原有 `[DynamicColumn]` 单字典和 `Mapping.DynamicColumns` 调用保持不变。

```csharp
var layout = ExcelEntity.LayoutFromAttributes<PurchaseOrder>(builder => builder
    .ListRegion("采购单", "A4", order => order.Lines, region => region
        .DynamicColumnGroup("goods", line => line.Goods, new[]
        {
            new ExcelDynamicColumnDefinition { Key = "brand", Title = "品牌" }
        })
        .DynamicColumnGroup("product", line => line.Product, new[]
        {
            new ExcelDynamicColumnDefinition { Key = "season", Title = "季节" }
        })
        .Footer("合计", footer => footer
            .Cell("B1", lines => lines.Sum(line => line.Quantity))
            .Cell("C1", lines => lines.Sum(line => line.Amount)))));
```

`Footer` 的相对 `A1` 是明细结束后的首行和列表起始列。它写入一个精确 marker，并可通过 `GapRows`、聚合单元格、相对合并和数字格式生成数量或金额尾部。导入时 marker 只用于确定明细边界，不会被物化成明细项；聚合值由业务模型或委托计算，不提供公式 DSL。固定单元格也可以使用 `[ExcelEntityCell("采购单", "B2")]` 与 `LayoutFromAttributes`，再通过 Fluent API 补充列表、合并和尾部配置。

## Entity Layout 连续分组小计

列表明细已经按业务顺序分组时，可以使用 `GroupSubtotal` 在相邻键变化处写入小计 Footer。它复用 Footer 的聚合单元格、相对合并、样式和 `GapRows`，不会自动排序，也不会把不连续的同名分组拼接在一起。

```csharp
var layout = ExcelEntity.Layout<PurchaseOrder>(builder => builder
    .ListRegion("明细", "A4", order => order.Lines, region => region
        .GroupSubtotal(line => line.Warehouse, "仓库小计", footer => footer
            .Cell("C1", lines => lines.Sum(line => line.Amount), numberFormat: "0.00"))
        .Footer("订单合计", footer => footer
            .Cell("C1", lines => lines.Sum(line => line.Amount), numberFormat: "0.00"))));
```

导出时每个连续分组先写明细，再写分组 marker 和小计内容；最终 `Footer` 仍写在全部分组之后。导入时会识别并跳过分组 marker、间隔行及其相对 Footer 内容，因此小计不会被转换为明细项。分组 marker 与最终 Footer marker 必须使用不同文本；分组键委托的异常和取消语义沿用当前导出调用。

## Entity Layout 水平分页

需要控制打印分页但不改变明细数据边界时，可以在列表区域配置 `PageBreak(rowsPerPage)`。`rowsPerPage` 只统计明细行，不包含表头；Provider 会在每个完整页面之后写入水平分页符，不插入空白行，也不会影响 Footer、marker 或导入结果。

```csharp
var layout = ExcelEntity.Layout<PurchaseOrder>(builder => builder
    .ListRegion("明细", "A4", order => order.Lines, region => region
        .PageBreak(20)
        .Footer("订单合计", footer => footer
            .Cell("C1", lines => lines.Sum(line => line.Amount), numberFormat: "0.00"))));
```

分页策略与 `GroupSubtotal` 互斥。分页小计和签字行可通过下一节的 `PageSubtotal` 与最终 `Footer` 配置；最终 `Footer` 可使用 `FooterNamed` 为总计 marker 创建或更新命名锚点。重复页小计及其签字区的逐页命名锚点尚未支持。NPOI 的 XLS/HSSF、XLSX/XSSF 与 ClosedXML XLSX 支持该导出能力，分页不会让 MiniExcel、SpreadCheetah 或 ExcelDataReader 获得 Entity Layout 能力声明。

## Entity Layout 分页小计

需要在每个完整分页明细段结束处输出小计和签字栏时，可以在 `PageBreak` 后配置 `PageSubtotal`。小计仅写在仍有下一段明细的完整页后；末段（包括不足一页的尾段）由普通 `Footer` 输出总计和签字栏。尾部的 `A1` 是 marker 保留位置，其他单元格坐标相对于 marker；例如 `B1`、`C1` 写小计或总计，`B2`、`C2` 写其下一行的签字栏。聚合委托接收当前页或全部明细的只读快照，导入时会按 marker、间隔和尾部高度跳过尾部区域。

```csharp
var layout = ExcelEntity.Layout<PurchaseOrder>(builder => builder
    .ListRegion("明细", "A4", order => order.Lines, region => region
        .PageBreak(20)
        .PageSubtotal("页小计", footer => footer
            .Cell("C1", lines => lines.Sum(line => line.Amount), numberFormat: "0.00")
            .Cell("B2", "经办人签字：________________")
            .Cell("C2", "复核人签字：________________"))
        .FooterNamed("OrderTotalMarker", "订单合计", footer => footer
            .Cell("C1", lines => lines.Sum(line => line.Amount), numberFormat: "0.00")
            .Cell("B2", "经办人签字：________________")
            .Cell("C2", "复核人签字：________________"))));
```

`FooterNamed` 会把工作簿定义名称 `OrderTotalMarker` 创建或更新为最终总计 marker 的实际单元格；模板中同一工作表已有的 global 或工作表局部单格名称会重定位到该 marker。导入仍以 marker 识别明细边界，分页中的 `PageSubtotal` 不会绑定这个名称。

`PageSubtotal` 必须同时配置 `PageBreak`，不能与 `GroupSubtotal` 组合；分页小计 marker 与最终 `Footer` marker 必须不同。`PageBreak(20)` 按每 20 条明细插入工作表水平分页符，不保证与打印机最终生成的物理页完全一致；纸张、页边距、缩放、重复标题、行高等打印设置可能改变实际分页，因此签字栏应结合目标打印设置检查。明细为空时不会产生分页小计，只会写入最终 `FooterNamed`，所以总计及末尾签字栏仍可出现。NPOI 的 XLS/HSSF、XLSX/XSSF 与 ClosedXML XLSX 支持该能力；导入不会把小计或签字行物化为明细。

尾部需要 Excel 公式时，使用显式的 `Formula`；不要依赖 `Cell` 接收以 `=` 开头字符串时的 Provider 旧行为：

```csharp
var layout = ExcelEntity.Layout<PurchaseOrder>(builder => builder
    .ListRegion("明细", "A4", order => order.Lines, region => region
        .Footer("订单合计", footer => footer
            .Formula("C1", "=SUM(C5:INDEX(C:C,ROW()-1))", numberFormat: "0.00"))));
```

`Formula` 接收以 `=` 开头的原生 Excel A1 公式，并将其写为公式单元格；`C1` 是相对尾部起点的地址。框架不计算公式结果，也不随明细行数改写公式引用，业务方需自行保证引用范围正确。格式错误的公式可能被底层工作簿引擎拒绝。NPOI XLS/XLSX 与 ClosedXML XLSX 均支持该方法。

当尾部上方是本次导出的连续明细，可以让构建器按实际明细行数生成求和公式：

```csharp
var layout = ExcelEntity.Layout<PurchaseOrder>(builder => builder
    .ListRegion("明细", "A4", order => order.Lines, region => region
        .Footer("订单合计", footer => footer
            .FormulaSumContiguousRowsAbove("C2", "B", numberFormat: "0.00")
            .GapRows(1))));
```

`B` 是工作表的绝对列字母，`C2` 相对于尾部 marker。方法会根据明细快照、最终 `GapRows` 和单元格相对行生成 `SUM(INDEX(...):INDEX(...))`；空明细写入 `=0`。`GapRows` 在公式方法前后配置均可。分页小计和分组小计各自收到连续明细，也可使用此方法；若最终 Footer 与中间小计同时存在，布局会拒绝在最终 Footer 使用它，以免把小计行重复计入。

最终 Footer 要跨分页或分组小计汇总明细时，在相应 `Footer` 或 `FooterNamed` 配置中使用 `FormulaSumDetailRowsAbove("C2", "B", numberFormat: "0.00")`。Provider 根据实际写出的明细行段生成 `SUM(B5:B6,B10:B11)` 等原生公式，小计、间隔和签字行都不进入引用范围；空明细写入 `=0`。该方法仅用于最终 Footer，不能用于 `PageSubtotal` 或 `GroupSubtotal`；公式不由本库计算。来源列必须是实际明细值所在的工作表列，布局的最终物理列位置应与业务定义核对。超过 Excel 公式长度限制时导出失败，已有目标文件不会被提交覆盖。NPOI XLS/XLSX 与 ClosedXML XLSX 均支持。

## Entity Layout 计算列

列表区域可以用 `CalculatedColumn` 声明序号、金额、差异等导出计算列。委托收到当前明细项、本次导出的只读快照、零基明细索引以及一基工作表行列号；每个导出单元格只执行一次。计算列支持 `Before`、`After` 或显式物理位置，也可以配置正文样式和数字格式。

```csharp
var layout = ExcelEntity.Layout<PurchaseOrder>(builder => builder
    .ListRegion("明细", "A4", order => order.Lines, region => region
        .CalculatedColumn<int>("line-no", "序号", row => row.Index + 1,
            column => column.Placement(ExcelColumnPlacement.Before("Code")))
        .CalculatedColumn<decimal>("amount", "金额",
            row => row.Item.Quantity * row.Item.UnitPrice,
            column => column.Placement(ExcelColumnPlacement.After("Quantity"))
                .NumberFormat("0.00"))));
```

计算列是导出专用定义。导入时仍保留其物理列位置，读取器跳过该列，不会尝试写入实体属性；计算委托不会在导入过程中执行。委托异常会返回包含 Sheet、行、列和 Key 的结构化导出异常，取消异常保持原样透传。

## Entity Layout 命名锚点

模板地址变化频繁时，可以用工作簿命名范围代替固定的 `A1` 地址。`CellNamed` 绑定固定字段，`ListRegionNamed` 绑定明细起点；两个 API 都要求显式提供布局工作表名称，名称在模板预检阶段必须解析为该工作表中的单个单元格。

```csharp
var layout = ExcelEntity.Layout<PurchaseOrder>(builder => builder
    .CellNamed("采购单", "OrderStatus", order => order.Status)
    .ListRegionNamed("明细", "OrderLines", order => order.Lines,
        region => region.Header()));
```

命名锚点支持工作簿级名称和对应工作表的局部名称；局部名称优先。缺少名称、名称指向多单元格区域、跨工作表或存在同作用域歧义时，Provider 在写入目标流前返回 `Plan` 配置错误。命名明细区域由运行时数据和 Footer 决定行边界，因此不能再配置绝对 `End` 地址；旧的固定地址 Fluent API 行为保持不变。

显式动态列组也可以直接声明内置校验规则。例如 `ValidationRules = new[] { new ExcelMappingDynamicValidationConfiguration { Name = "range", Min = 1, Max = 10 } }`；支持 `required`、`regex`、`date`、`maxValue`、`range`、`maxLength` 和 `unique`。规则在 NPOI 与 ClosedXML 中复用同一 Mapping Plan，并返回统一的行列坐标错误。

实体导入默认在当前行首次校验失败后停止继续处理该行。需要收集同一行的全部校验错误时，可以传入 `new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue)`；该选项只改变当前行的校验遍历，不改变资源限制、取消和文件流所有权。

CSV 也使用动态属性绑定：

```csharp
using var input = File.OpenRead("orders.csv");
using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
var importer = provider.GetRequiredService<ICsvImporter>();
var result = importer.Import<OrderRow>(input);
var region = result.Items[0].Values["区域"];
```
