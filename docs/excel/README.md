# Excel 高级导入导出

## Excel 主路径

导入和导出统一使用 Workbook Request。Mapping 的最终优先级为 `Attribute < Profile < JSON/XML Document < Request Fluent`，仅 Plan compiler 完成最终合并；JSON/XML 保留为 Document 级 API。

导出模板写入默认使用 `PreserveTemplate`，保留目标单元格的模板样式和批注；需要清除模板样式和批注时，显式选择 `ReplaceTemplate`。NPOI 管线是 DOM 管线，不承诺 streaming 或零 GC。

Workbook 元数据使用请求级 `ExcelWorkbookMetadataOptions` 配置，不依赖进程级默认状态。未显式调用 `Metadata(...)` 时模板 metadata 保留；显式调用时六个字段覆盖模板值，XLS 与 XLSX 采用相同策略。`IExcelExporter.ExportToFile` 和 `ICsvExporter.ExportToFile` 先写同目录临时文件，成功关闭并 flush 后在最终提交点再次检查取消，再替换目标；导出失败或取消不会截断已有目标文件。直接写入调用方 Stream 时不提供回滚保证，失败后可能已经产生部分写入。

2.0.0 已移除与上述接口实例方法重复的四个 `ExportToFile`/`ExportToFileAsync` 扩展声明。请使用 `ICsvExporter` 或 `IExcelExporter` 实例调用同名成员；完整迁移示例见 [nuget-migration.md](nuget-migration.md)。

NPOI 导入会先把输入复制到受 `MaxInputBytes` 约束的内存流，再建立 Workbook DOM；`ImportAsync` 的输入复制使用 `ReadAsync`，之后的 NPOI DOM 解析仍为同步阶段；`MaxInputBytes` 不等于解压后 DOM 峰值保护。失败工作簿的 `MaxSerializedBytes` 只限制序列化输出，不限制原始 Workbook、解压内容、业务实体或失败工作簿 DOM 的峰值。部署不受信任文件时还应设置进程内存/CPU 限额，并使用 `ExcelResourceLimits` 限制行、错误、图片和唯一值；未映射图片不会被图片限制器扫描。

MiniExcel 是与 NPOI 并列的独立 XLSX Provider。它复用同一套 Workbook Request、Mapping Plan、Converter、Validation、错误结果和文件提交契约，不引入第二套 `MiniExcel*Request`、Mapping Profile 或业务 Attribute。通过 `services.AddBingOfficesMiniExcel()` 注册后，业务代码仍只依赖 `IExcelImporter` 与 `IExcelExporter`。应用启动时应选择一个 Provider；两个注册扩展都使用 `TryAdd`，同时注册时先注册的实现会保留，不能把这种顺序行为当作按请求动态切换 Provider 的 API。

下方完整 Workbook 能力列表以 NPOI Provider 为基准；MiniExcel 只承诺 [Provider 能力矩阵](09-providers.md) 中已验证的常规 XLSX、映射、转换、校验、动态列和关系能力，其余请求必须按结构化 `UnsupportedFeature` fail-fast。

生产 API 以 Workbook Request 为中心，支持：

- 异构多 Sheet 导出和导入
- typed 动态列、别名、稳定 Key、类型转换和物理布局
- 模板 Workbook、命名区域、样式缓存和自定义表头
- 显式父子关系绑定与结构化导入错误
- XLSX 柱状、折线和饼图

### Sync 与 Async

`IExcelImporter.ImportAsync`、`IExcelExporter.ExportAsync` 和 `ExportToFileAsync` 仅将外围文件/Stream IO 异步化；NPOI DOM 构建、映射、校验和序列化仍运行在调用线程。模板流在 `WorkbookFactory.Create` 阶段仍要求同步读取，因此 async-only 模板流不属于支持的输入合同。CSV 的 `ImportAsync`、`ExportAsync` 和 `ExportToFileAsync` 使用异步 CSV Parser/Writer 与 Stream IO。两套 API 共享 Mapping、Validation、Conversion、错误结果和取消合同，不应使用 `Task.Run` 替代真实 IO。

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
- [exceptions-and-observers.md](exceptions-and-observers.md)：异常分类、Observer 与文件提交边界
- [dates.md](dates.md)：日期、DateTimeOffset 与跨时区合同
- [npoi-extensions.md](npoi-extensions.md)：七个 NPOI 用户扩展容器和 Try/Throw 行为
- [async-io.md](async-io.md)：Sync/Async API、流所有权、取消和 NPOI 同步边界
- [09-providers.md](09-providers.md)：NPOI 与 MiniExcel Provider 的能力矩阵、选择和资源边界
- [nuget-migration.md](nuget-migration.md)：包身份、当前兼容边界和迁移注意事项

迁移与使用：

- [mapping-profile.md](mapping-profile.md)
- [mapping-json-xml.md](mapping-json-xml.md)
- [import-validation.md](import-validation.md)
- [dynamic-columns.md](dynamic-columns.md)
- [exceptions-and-observers.md](exceptions-and-observers.md)
- [dates.md](dates.md)
- [npoi-extensions.md](npoi-extensions.md)
- [async-io.md](async-io.md)
- [09-providers.md](09-providers.md)
- [nuget-migration.md](nuget-migration.md)

ASP.NET Core 上传示例见 `import-validation.md`；公开示例由 `Bing.Offices.Docs.Tests` 对当前源码持续编译和执行，真实 nupkg 可消费性由发布任务中的独立 PackageConsumer 验证。
