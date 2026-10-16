# 工作表保护与填写区域

NPOI 的 XLS、XLSX 和 ClosedXML 的 XLSX 导出支持工作表保护。保护只限制 Excel 界面中的误编辑，不提供文件加密或访问控制；导入时仍应执行服务端校验。

下面的填写模板锁定普通单元格，仅允许编辑第 2 至 101 行的 A 列。区域使用零基闭区间坐标，可以覆盖尚未写入内容的空白单元格。

```csharp
var protection = new ExcelSheetProtectionDefinition(new[]
{
    new ExcelRangeDefinition
    {
        StartRow = 1, EndRow = 100,
        StartColumn = 0, EndColumn = 0
    }
}, password: "template-password");

var request = ExcelExport.Workbook(workbook => workbook
    .AddSheet("采购单", orders, sheet => sheet.Protection(protection)));
exporter.ExportToFile(request, path);
```

Entity Layout 使用相同的定义：

```csharp
var layout = ExcelEntity.Layout<PurchaseOrder>(builder => builder
    .Cell("采购单", "B2", order => order.Number)
    .Protection("采购单", protection));
```

密码可省略，此时仍启用无密码工作表保护。传入空的区域列表会保护整个工作表。配置保护时只解锁指定区域，其他单元格沿用新工作簿或模板中原有的锁定状态；已有模板样式的颜色、边框和数字格式不会被替换。未配置保护时，模板原有保护保持不变。保护区域不随明细行数自动移动，调用方应按模板实际填写区声明；XLS/XLSX 的行列上限在写入目标前校验。

MiniExcel 和 SpreadCheetah 不支持此能力，遇到保护请求会在写入前返回结构化 `UnsupportedFeature`。导出器的 `IExcelProviderFeatureDescriptor` 可用 `ExcelProviderFeatures.SheetProtection` 查询该能力。旧的 Workbook 与 Entity 调用无需迁移；只有需要模板填写区保护时才添加上述配置。

## 核心实现与依赖配置

实现为 `src/Bing.Offices.Abstractions/Bing/Offices/Exports/Content/ExcelSheetProtectionDefinition.cs` 及 NPOI/ClosedXML 的 `Bing/Offices/Exports/*SheetProtectionWriter.cs`。先校验区域和格式上限，再为解锁区复制/调整锁定样式并启用保护。依赖已注册的数据 Provider，保护定义随 Workbook/Entity 请求传入，没有独立加密配置；模板保护密码与商业扩展的文件加密密码属于不同能力。
