# 动态列与实体分页布局

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

### Entity Layout 多个隐式动态字典

Forward 批次导入可以让同一个明细模型的多个 `[DynamicColumn]` 字典分别接收动态表头。单个字典仍可沿用旧配置；多个字典时，`ExcelMappingDynamicColumnConfiguration.TargetPropertyName` 是必填归属信息，必须指向明细类型上的直接可写 `IDictionary<string, object>` 属性：

```csharp
var mapping = new ExcelMappingConfiguration
{
    Columns = new List<ExcelColumnConfiguration>
    {
        new ExcelColumnConfiguration { PropertyName = "Code", Title = "编码" }
    },
    DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
    {
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "zone",
            Title = "区域",
            TargetPropertyName = "Attributes",
            ColumnIndex = 1
        },
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "note",
            Title = "备注",
            TargetPropertyName = "Extensions",
            ColumnIndex = 2
        }
    }
};
```

`TargetPropertyName` 在计划阶段参与属性、标题、Key 和物理列冲突校验；同一 Key 不能归属多个字典，缺少归属或指向不可写属性会返回配置错误。NPOI XLS/XLSX 与 ClosedXML XLSX 的 Entity Layout 完整导出/导入，以及 NPOI/ClosedXML 的 Entity Forward 同步/异步批次，都会把值写回对应字典，并保留转换器、未知动态列策略、取消和原始坐标语义。NPOI HSSF Forward 仅支持无模板、单 Sheet、单列表的受限路径；同一工作表的多个列表仍沿用各自已有边界，需要复杂布局时使用完整入口。

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

分页策略可与 `GroupSubtotal` 组合；组跨页时继续累计，组结束后再写分组小计。如果页末恰好也是组末，分页符位于分组小计之后。分页小计和签字行可通过下一节的 `PageSubtotal` 与最终 `Footer` 配置；`PageSubtotalNamed` 可为各中间页小计 marker 创建顺序命名锚点，最终 `Footer` 可使用 `FooterNamed` 为总计 marker 创建或更新命名锚点。NPOI 的 XLS/HSSF、XLSX/XSSF 与 ClosedXML XLSX 支持该导出能力，分页不会让 MiniExcel、SpreadCheetah 或 ExcelDataReader 获得 Entity Layout 能力声明。

### 重复打印标题与实际页眉

只需打印时在每张纸上重复已有标题行，可在 Entity 布局配置 `PrintTitleRows`。`1:3` 使用工作表的一基行号；它设置打印元数据，不插入单元格，也不要求 `PageBreak`：

    var printLayout = ExcelEntity.Layout<PurchaseOrder>(builder => builder
        .PrintTitleRows("明细", "1:3")
        .ListRegion("明细", "B3", order => order.Lines));

需要导出文件中每个后续逻辑页实际拥有页眉行时，改用 `PageHeader`。下面假定模板的 `A1:D3` 是首 页页眉，列表表头从 `B3` 开始；`A1` 是专供导入识别的 marker。无模板时可用 `Cell` 和 `Merge` 填充首 页区域：

    var layout = ExcelEntity.Layout<PurchaseOrder>(builder => builder
        .ListRegion("明细", "B3", order => order.Lines, region => region
            .PageBreak(20)
            .PageHeader("A1:D3", "A1", "页眉", header => header
                .Cell("D1", page => $"第 {page.PageNumber}/{page.PageCount} 页")
                .Cell("A2", page => $"本页 {page.Items.Count} 行"))
            .PageSubtotalNamed("PageTotal", "页小计")
            .Footer("订单合计")));

页眉 `Cell`、`Merge` 的坐标相对于源区域左上角，`markerAddress` 则是工作表绝对坐标。首 页先写 marker 和覆盖值；后续页复制源区域的普通单元格值、公式文本、样式（含锁定状态）、行高、隐藏状态及完整包含的合并区域，再计算每页覆盖值。源区域包含富文本、超链接、批注、数据校验或条件格式时，多页导出在写目标前拒绝，避免后续页静默丢失格式或交互规则；布局声明的数据校验若覆盖源区域也会拒绝。单页及空明细无需复制，原模板元数据保留。启用列表表头时，`Cell` 可为标题单元格设置样式或写入相同标题，但不能改写最终映射标题；`Merge` 不得触及列表表头。导入按 marker 和 `PageBreak` 容量跳过复制页眉，并核对复制页的列表表头，不把它当作明细。空明细或单页不插入副本；同一边界先写分组小计、再写分页小计，分页符之后才写下一页页眉。`PageSubtotalNamed` 始终指向小计 marker。

同一 Sheet 不能同时配置 `PrintTitleRows` 与 `PageHeader`；模板已有的打印标题也会与实际页眉冲突。同一 Sheet 至多配置一个实际重复页眉区域，可同时配置不重叠的普通列表区域。源区域末行须紧邻首个明细行并包含启用的完整列表表头；使用 `ListRegionNamed` 时，导入和导出都会在解析名称后检查此关系。marker 不得占用固定字段或表头，也不得与小计和最终 Footer 使用相同文本。marker 文本是其所在列的保留值，明细写出相同文本会在输出前报配置错误；无列表表头时，marker 列必须位于明细物理列范围外。后续页眉目标区域不能占用已声明的固定字段（即使值为空），也不能已有内容或冲突合并。`End` 可以预留空白填写行，导入以区域内实际内容确定明细终点。页码是 `PageBreak` 定义的逻辑页码，纸张和缩放可能使最终打印页数不同。公式只复制文本，不执行计算；图片、图表等绘图对象不会复制。

重复页眉导入在末行回溯、marker 扫描及复制页表头校验期间响应取消。稀疏工作表中的远端非列表单元格不会被当作明细；显式的大 `End` 仍只约束区域边界。输入字节上限在打开工作簿时检查，明细行数上限在读取明细时检查。

### 独立逐页页眉与签字区名称

需要让外部模板工具定位每个逻辑页的页眉和签字区时，在已有 `PageHeader` 或每页签字 Footer 上使用独立名称入口。两个入口都必须先配置 `PageBreak`；页容量只统计明细行，空明细也视为一个逻辑页。首页、末页、单页和空页都会生成对应的 `Head_001`、`Sign_001` 等单格名称，名称指向实际 marker，不表示 Excel 打印机最终产生的物理页。

```csharp
var layout = ExcelEntity.Layout<PurchaseOrder>(builder => builder
    .ListRegion("明细", "B3", order => order.Lines, region => region
        .PageBreak(20)
        .PageHeaderNamed("Head", "A1:D3", "A1", "页眉", header =>
            header.Cell("D1", page => $"第 {page.PageNumber}/{page.PageCount} 页"))
        .PageFooterNamed("Sign", "签字", footer => footer
            .GapRows(1)
            .Cell("B1", lines => lines.Count)
            .Cell("C1", lines => lines.Sum(line => line.Amount))
            .Cell("A2", "审核"))
        .Footer("订单合计", footer => footer
            .Cell("C1", lines => lines.Sum(line => line.Amount)))));
```

`PageFooter` 可以省略名称，只输出每页签字区；`PageFooterNamed` 在此基础上为每页 marker 创建顺序名称。每页 Footer 位于该页明细以及同边界分组/分页小计之后、分页符和下一页页眉之前，末页 Footer 位于最终 Footer 之前。Footer 委托只接收当前页明细的只读快照，样式、数字格式、合并和 `GapRows` 沿用普通 Footer。

`PageHeaderNamed` 只为已经存在的实际页眉命名，不改变 `PageHeader` 的复制、`PrintTitleRows` 互斥和 marker 校验行为。真实打印页数仍受纸张、缩放、行高和重复标题影响；逻辑页名称不能替代打印渲染验证。导入只扫描 marker 所在列，并按页容量、尾部高度和间隔跳过页眉/签字区，不把签字数字转换为明细。

名称策略默认为 `ExcelEntityPageAnchorStaleNamePolicy.Reject`。模板中的跨 Sheet、二维范围、公式、失效引用或同名多作用域旧名称即使选择 `RemoveObsolete` 也会在 `Plan` 阶段拒绝；只有属于当前系列、目标 Sheet 上合法单格且本次未使用的名称才会被清理。清理只发生在输出副本，输入模板和失败前的旧输出保持不变。最终 Footer 使用 `FormulaSumDetailRowsAbove` 时，Provider 只把实际明细行段写入公式区间，分页/分组小计、空白行和每页签字数字均被排除；公式不由库计算。

当前实现范围为 NPOI 的 XLS/HSSF、XLSX/XSSF 和 ClosedXML 的 XLSX；MiniExcel、SpreadCheetah 与 ExcelDataReader 不因此获得 Entity Layout 能力声明。

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

需要让公式或外部程序逐页定位中间小计时，使用 `PageSubtotalNamed("PageTotal", "页小计", ...)` 替代 `PageSubtotal`。每个有下一页的完整明细页依次生成 `PageTotal_001`、`PageTotal_002` 等名称，指向该页小计 marker；末页、单页和空明细不生成分页名称。模板中同名、同 Sheet 的单格名称会随本次布局重定位并保留原作用域；若页数缩短后模板仍含同系列旧名称，或名称引用其他 Sheet、范围、公式、存在歧义，导出在写入目标前以 `Plan` 错误拒绝。导入仍以 marker 为准。

`PageSubtotal`、`PageSubtotalNamed` 都必须同时配置 `PageBreak`；分页、分组和最终尾部的 marker 必须各不相同。与 `GroupSubtotal` 组合时，页容量只计算明细；页聚合接收当前页明细，组聚合接收完整连续组。同一明细行同时结束组和非末页时，先写分组小计、再写分页小计，分页符位于两者之后。导入按 marker 跳过两种小计并校验每个中间页的明细数量。`PageBreak(20)` 不保证与打印机最终生成的物理页完全一致；纸张、页边距、缩放、重复标题、行高等打印设置可能改变实际分页。明细为空时只写最终 `FooterNamed`。NPOI 的 XLS/HSSF、XLSX/XSSF 与 ClosedXML XLSX 支持该能力。

```csharp
var groupedPages = ExcelEntity.Layout<PurchaseOrder>(builder => builder
    .ListRegion("明细", "A4", order => order.Lines, region => region
        .GroupSubtotal(line => line.Warehouse, "仓库小计", footer => footer
            .Cell("C1", lines => lines.Sum(line => line.Amount)))
        .PageBreak(20)
        .PageSubtotal("页小计", footer => footer
            .Cell("C1", lines => lines.Sum(line => line.Amount)))
        .Footer("订单合计", footer => footer
            .Cell("C1", lines => lines.Sum(line => line.Amount)))));
```

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

`B` 是工作表的绝对列字母，`C2` 相对于尾部 marker。方法会根据明细快照、最终 `GapRows` 和单元格相对行生成 `SUM(INDEX(...):INDEX(...))`；空明细写入 `=0`。`GapRows` 在公式方法前后配置均可。单独配置的分页或分组小计可使用此方法；两种中间小计组合时，布局会拒绝连续求和公式，因为另一种小计可能插入到求和区间。最终 Footer 跨中间小计时也不能使用它。

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

## 核心实现、依赖与配置

公共动态列定义位于 `src/Bing.Offices.Abstractions/Bing/Offices/Exports/DynamicColumns`；Entity 分组、计算列、页眉和 Footer 在同项目 `Entities/Lists`。Core 的 `Mappings` 编译最终列位置、转换与校验；`src/provider-shared/ExcelEntityRowPlan.cs` 规划明细、小计和页眉/尾部的物理行，NPOI/ClosedXML 的 `Entities` 执行实际读写。

使用前注册所选 Provider 和命名转换/校验服务；Entity 能力仅 NPOI/ClosedXML。地址、marker、动态列 Key、组标识和物理位置是不同概念，模板布局变化后应重新预检。分批摘要不是完整布局的任意低内存替代，另见[批量导入](batch-import.md)；固定字段与模板基础用法见[实体布局](entity-layout.md)。
