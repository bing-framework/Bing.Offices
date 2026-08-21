# Excel 高级导入导出

生产 API 以 Workbook Request 为中心，支持：

- 异构多 Sheet 导出和导入
- typed 动态列、别名、稳定 Key、类型转换和物理布局
- 模板 Workbook、命名区域、样式缓存和自定义表头
- 显式父子关系绑定与结构化导入错误
- XLSX 柱状、折线和饼图

```csharp
var request = ExcelExport.Workbook(workbook =>
    workbook.AddSheet("订单", orders));
exporter.Export(request, stream);
```

详细设计见 `ai_docs/excel/00-overview.md` 至 `08-validation.md`。
