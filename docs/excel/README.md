# Excel 高级导入导出

## Excel 主路径

导入和导出统一使用 Workbook Request。Mapping 的最终优先级为 `Attribute < Profile < JSON/XML Document < Request Fluent`，仅 Plan compiler 完成最终合并；JSON/XML 保留为 Document 级 API。

导出模板写入默认使用 `PreserveTemplate`，保留目标单元格的模板样式和批注；需要清除模板样式和批注时，显式选择 `ReplaceTemplate`。NPOI 管线是 DOM 管线，不承诺 streaming 或零 GC。

Workbook 元数据使用请求级 `ExcelWorkbookMetadataOptions` 配置，不依赖进程级默认状态。未显式调用 `Metadata(...)` 时模板 metadata 保留；显式调用时六个字段覆盖模板值，XLS 与 XLSX 采用相同策略。`ExportToFile` 和 CSV `ExportToFile` 先写同目录临时文件，成功关闭并 flush 后在最终提交点再次检查取消，再替换目标；导出失败或取消不会截断已有目标文件。直接写入调用方 Stream 时不提供回滚保证，失败后可能已经产生部分写入。

NPOI 导入会先把输入复制到受 `MaxInputBytes` 约束的内存流，再建立 Workbook DOM；`MaxInputBytes` 不等于解压后 DOM 峰值保护。部署不受信任文件时还应设置进程内存/CPU 限额，并使用 `ExcelResourceLimits` 限制行、错误、图片和唯一值；未映射图片不会被图片限制器扫描。

生产 API 以 Workbook Request 为中心，支持：

- 异构多 Sheet 导出和导入
- typed 动态列、别名、稳定 Key、类型转换和物理布局
- 模板 Workbook、命名区域、样式缓存和自定义表头
- 显式父子关系绑定与结构化导入错误
- XLSX 柱状、折线和饼图

```csharp
var request = ExcelExport.Workbook(workbook =>
    workbook.Metadata(new ExcelWorkbookMetadataOptions { Author = "业务系统" })
        .AddSheet("订单", orders));
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
