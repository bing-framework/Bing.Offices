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

导入时，如果两个不同表头（正式标题、属性名或别名）绑定到同一个固定属性或已声明动态 Key，会返回结构化 `InvalidHeader` 错误，不会由后列表头静默覆盖前列数据。CSV 使用 `HasHeader = false` 导入动态字典属性时，必须声明动态列定义；列按映射属性顺序展开，动态列再按 `Order`、`Key` 排序，与导出顺序一致。

CSV 也使用动态属性绑定：

```csharp
using var input = File.OpenRead("orders.csv");
using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
var importer = provider.GetRequiredService<ICsvImporter>();
var result = importer.Import<OrderRow>(input);
var region = result.Items[0].Values["区域"];
```
