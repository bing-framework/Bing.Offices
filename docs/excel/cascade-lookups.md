# 两级联动字典模板

`CascadeList` 为上级字段生成文本下拉，并按上级值为同一行的下级字段生成候选。Workbook 与 Entity Layout 均可配置；NPOI 支持 XLS/XLSX，ClosedXML 支持 XLSX。字典在构建请求时快照化，输出为隐藏 Sheet、工作簿名称和原生列表校验。辅助 Sheet 追加在业务 Sheet 后。

```csharp
var products = new Dictionary<string, IReadOnlyList<string>>
{
    ["电子"] = new[] { "手机", "电脑" },
    ["食品"] = new[] { "饮料", "零食" }
};

var request = ExcelExport.Workbook(book => book
    .CascadeList("ProductCategories", products)
    .AddSheet("商品", rows, sheet => sheet
        .DataValidation(new ExcelDataValidationDefinition
        {
            Range = new ExcelRangeDefinition
            {
                StartRow = 1, EndRow = 500, StartColumn = 0, EndColumn = 0
            },
            Type = ExcelDataValidationType.List,
            ListSourceName = "ProductCategories"
        })
        .DataValidation(new ExcelDataValidationDefinition
        {
            Range = new ExcelRangeDefinition
            {
                StartRow = 1, EndRow = 500, StartColumn = 1, EndColumn = 1
            },
            Type = ExcelDataValidationType.List,
            CascadeSourceName = "ProductCategories",
            ParentColumnIndex = 0
        })));
```

Entity 布局使用相同的 `CascadeList`，以 `DataValidation("商品", definition)` 配置固定坐标。`ParentColumnIndex` 是零基列号；上级列必须位于下级单列区域以外，引用始终在同一行。多个上级字典可分别声明。上级文本按大小写不敏感规则唯一，下级文本保留逗号、引号、换行及以 `=` 开头的字面内容。

下级区域写入 `INDIRECT(VLOOKUP(...))` 公式，供 Excel 填写时显示候选。未匹配的上级值不提供下级候选。本功能不计算公式；以 `WorkbookRules` 或 `ConfiguredAndWorkbook` 导入这类单元格时，原生公式按现有 `ExcelUnsupportedFeaturePolicy` 处理。服务端仍需通过映射、转换器及配置规则校验上级和下级组合，不能将工作簿下拉当作业务授权规则。MiniExcel 与 SpreadCheetah 对级联导出请求在写入前返回 Unsupported。

## 核心实现与依赖配置

公共定义位于 `src/Bing.Offices.Abstractions/Bing/Offices/Exports/Content/ExcelCascadeListDefinition.cs`；NPOI/ClosedXML 的 `Bing/Offices/Exports/*CascadeListWriter.cs` 负责隐藏表、名称映射和列表公式。依赖各自 Provider；字典由调用方传入，库不查询外部字典服务。模板或透视占用辅助表名时按实际名称避让，业务端不要硬编码生成的隐藏 Sheet 名称。
