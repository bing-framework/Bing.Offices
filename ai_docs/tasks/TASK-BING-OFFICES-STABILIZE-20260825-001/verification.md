# 验证记录

## 已执行并通过

| 范围 | 命令/过滤器 | 结果 |
| --- | --- | --- |
| ValidationMode | `Import_ValidationMode_ShouldSelectConfiguredAndWorkbookRules` | 4/4 passed |
| DI 默认规则 | `AddNpoi_DefaultValidationRules_ShouldMatchDirectConstruction` | passed |
| 失败工作簿 | `Import_ErrorRowsOnly_ByIndex_ShouldUseResolvedSheetHeader` | 1/1 passed |
| 模板既有回归 | `Export_TemplateRegion`, `Export_RequestStyle`, `Export_HeaderComment` | 5/5 passed |
| 模板覆盖策略 | `Export_TemplateCellOverwrite` | 2/2 passed |
| NPOI 生产构建 | 由上述测试触发的 Release 构建 | passed，存在既有兼容性/过时 API/XML 注释警告 |
| Unit net8 全量 | `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net8.0 -c Release --no-restore` | 213/213 passed |
| Integration net8 全量 | `dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net8.0 -c Release --no-restore` | 11/11 passed |
| Integration net6 全量 | `dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net6.0 -c Release --no-restore` | 11/11 passed |
| Docs consumer net8 | `dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -f net8.0 -c Release --no-restore` | 8/8 passed |
| Profile/API 聚焦 | `MappingProfileRegistryTest` + `PublicApiContractTest` | 17/17 passed |
| Benchmark Dry smoke | `dotnet run -c Release --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -- --filter *MappingValidationBenchmarks* --job Dry` | 208 benchmarks executed, no runtime failure |
| Benchmark 编译 | `dotnet build benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -f net8.0 -c Release --no-restore` | passed |
| Unit net6 全量 | `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net6.0 -c Release --no-restore` | 213/213 passed |
| 解决方案 Release build | `dotnet build Bing.Offices.sln -c Release --no-restore` | passed |
| 解决方案 pack | `dotnet pack Bing.Offices.sln -c Release --no-build --no-restore` | passed；未发布 |
| ResourceProbe | `dotnet run -c Release --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -- --resource-probe artifacts/resource-probe-task-20260825.jsonl` | 16/16 scenarios passed |
| 流协作者重构 Unit | `StreamPipelineTest` + `ExcelWorkbookRequestTest` + `ExcelP0RegressionTest` | 133/133 passed |
| 流协作者重构 Integration | `ExcelImporterIntegrationTest` | 11/11 passed |
| NPOI 流协作者构建 | `dotnet build src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj -f net8.0 -c Release --no-restore` | passed |
| Profile Core 迁移 API | `MappingProfileRegistryTest` + `PublicApiContractTest` | 17/17 passed；Profile 注册入口位于 Core，NPOI 仅保留 `AddNpoi` |
| Profile Core 迁移 Docs | `dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -f net8.0 -c Release --no-restore` | 8/8 passed |
| NPOI 流协作者直接测试 | `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net8.0 -c Release --no-restore --filter "FullyQualifiedName~StreamPipelineTest"` | 79/79 passed |
| 失败工作簿 writer 拆分 Unit | net8.0 Release 测试，过滤 `ExcelP0RegressionTest` 与 `StreamPipelineTest` | 104/104 passed |
| 失败工作簿 writer 拆分 Integration | `dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net8.0 -c Release --no-restore --filter "FullyQualifiedName~ExcelImporterIntegrationTest"` | 11/11 passed |
| CSV 表头绑定拆分 | `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net8.0 -c Release --no-build --no-restore --filter "FullyQualifiedName~CsvTest"` | 24/24 passed |
| CSV 动态类型解析拆分 | net8 CSV 聚焦 24/24；net6 CSV 聚焦 24/24；net8 Integration 11/11 | 导入/导出共用受限类型白名单，回归通过 |
| CSV 属性绑定文件拆分 | `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net8.0 -c Release --no-build --no-restore` | 213/213 passed |
| NPOI 失败 writer 构建 | `dotnet build src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj -f net8.0 -c Release --no-restore` | passed，4 个既有 obsolete warning |
| 固定 Mapping 性能基线 | `dotnet run -c Release --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -- --filter "*MappingValidationBenchmarks.DynamicPlanBuildCacheHit*"` | 16 个参数组合完成；固定 1 launch / 2 warmup / 3 measurement，报告 Mean、Allocated、Gen0/Gen1 |
| NuGet 包元数据审计 | 重新 `dotnet pack` 后检查三个 2.0.0 `.nupkg/.snupkg` | 三包均含 `README.md`、`LICENSE`、`icon.png`；nuspec 含 README 元数据；SourceLink 不在运行时 dependencies |

## 证据边界

- 本记录只包含本轮实际执行结果，不引用历史进度文档作为当前证据。
- net5.0、net7.0、netcoreapp3.1 Unit 仅完成编译，运行阶段分别缺少 .NET 5、.NET 7、.NET Core 3.1 runtime；未安装运行时或绕过 testhost。
- Excel/LibreOffice/WPS 互操作尚未执行；当前验证为 NPOI reopen 和 xUnit 回归。
- Profile 扫描部分加载容错的实现已存在，但尚未完成真正构造 `ReflectionTypeLoadException` 的直接测试；扫描 alias 元数据也尚未实现。
- 未执行 Excel/LibreOffice/WPS 互操作；net5/net7/netcoreapp3.1 Unit 仅编译，因缺少对应 runtime 无法运行；未执行完整 Stream import/export 性能前后对比或本地包安装消费者链路。

## 生产符号到测试映射

| 生产符号/行为 | 测试项目 | 测试方法或范围 |
| --- | --- | --- |
| `NpoiExcelImporter` ValidationMode 分流 | `Bing.Offices.Tests` | `ExcelP0RegressionTest.Import_ValidationMode_ShouldSelectConfiguredAndWorkbookRules` |
| `NpoiExcelImporter` 失败 Sheet identity | `Bing.Offices.Tests` | `ExcelP0RegressionTest.Import_ErrorRowsOnly_ByIndex_ShouldUseResolvedSheetHeader` |
| `ExcelTemplateCellOverwritePolicy` | `Bing.Offices.Tests` | `ExcelWorkbookRequestTest.Export_TemplateCellOverwrite*` |
| `ExcelMappingPlanFactory` 方向和 Patch | `Bing.Offices.Tests` | `MappingConfigurationPatchTest`、`ExcelWorkbookRequestTest.ColumnPlan_FixedAndDynamic_ShouldShareCompiledExecutionMetadata` |
| `IMappingProfileResolver`/稳定 Profile 名称 | `Bing.Offices.Tests` | `MappingProfileRegistryTest.ExplicitRegistration_WithStableName_ShouldResolveThroughReadOnlyResolver` |
| JSON/XML v2 安全和 round-trip | `Bing.Offices.Tests` | `ReviewFixRegressionTest`、`MappingConfigurationPatchTest` |
| `NpoiStreamCopier` 接入 | Unit/Integration | `StreamPipelineTest`、`ExcelWorkbookRequestTest`、`ExcelImporterIntegrationTest` |
| `NpoiStreamCopier` 复制/限制/取消/ownership | `Bing.Offices.Tests` | `StreamPipelineTest.NpoiStreamCopier_Copy_ShouldPreserveDataAndStreamOwnership`、`NpoiStreamCopier_Copy_WhenMaxBytesExceeded_ShouldThrowBeforeWrite`、`NpoiStreamCopier_Copy_WhenCancelled_ShouldThrowBeforeIo` |
| `NpoiFailureWorkbookWriter` 失败摘要与 ErrorRowsOnly | `Bing.Offices.Tests` / Integration | `ExcelP0RegressionTest`、`ExcelImporterIntegrationTest` 失败工作簿相关回归范围 |
