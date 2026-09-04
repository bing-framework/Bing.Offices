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

CSV 也使用动态属性绑定：

```csharp
using var input = File.OpenRead("orders.csv");
using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
var importer = provider.GetRequiredService<ICsvImporter>();
var result = importer.Import<OrderRow>(input);
var region = result.Items[0].Values["区域"];
```
