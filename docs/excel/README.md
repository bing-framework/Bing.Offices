# Excel 高级导入导出

## Excel 主路径

导入和导出统一使用 Workbook Request。Mapping 的最终优先级为 `Attribute < Profile < JSON/XML Document < Request Fluent`，仅 Plan compiler 完成最终合并；JSON/XML 保留为 Document 级 API。

导出模板写入默认使用 `PreserveTemplate`，保留目标单元格的模板样式和批注；需要清除模板样式和批注时，显式选择 `ReplaceTemplate`。NPOI 管线是 DOM 管线，不承诺 streaming 或零 GC。

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

按主题阅读：

- [mapping-profile.md](mapping-profile.md)：Profile 与映射优先级
- [mapping-json-xml.md](mapping-json-xml.md)：JSON/XML 映射文档
- [import-validation.md](import-validation.md)：Workbook 原生校验、配置校验和错误收集
- [dynamic-columns.md](dynamic-columns.md)：动态列与物理布局
- [nuget-migration.md](nuget-migration.md)：包身份、当前兼容边界和迁移注意事项

迁移与使用：

- [mapping-profile.md](mapping-profile.md)
- [mapping-json-xml.md](mapping-json-xml.md)
- [import-validation.md](import-validation.md)
- [dynamic-columns.md](dynamic-columns.md)
- [nuget-migration.md](nuget-migration.md)

ASP.NET Core 上传示例见 `import-validation.md`；公开示例由 `Bing.Offices.Docs.Tests` 使用本地打包的三个 NuGet 包持续编译和执行。
