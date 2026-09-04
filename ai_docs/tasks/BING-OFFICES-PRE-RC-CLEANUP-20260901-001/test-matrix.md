# 测试矩阵

| 生产符号/行为 | Unit 项目与方法 | Integration/Docs/Consumer | 当前状态 |
| --- | --- | --- | --- |
| `ExcelMappingDocumentFactory` 请求配置 | `Bing.Offices.Tests.ReviewFixRegressionTest.MappingDocumentFactory_RequestConfiguration_ShouldMergeOnlySelectedDirection`；`Bing.Offices.Tests.StreamPipelineTest.TypeMap_NormalizedDocument_ShouldCompileSelectedDirection` | Docs/consumer 使用方向化 Document | VERIFIED（net6/net8 Unit 中执行） |
| `ExcelMappingConfigurationLoader.MigrateV1Json/Xml` | `ReviewFixRegressionTest.MigrationV1_ShouldKeepNonTargetDirectionNull`；`StreamPipelineTest.MappingConfigurationLoader_Utf8FilesAndStreams_ShouldKeepCallerStreamsOpen` 等 | Docs `MappingDocuments_ExternalConsumer_ShouldMigrateAndPreserveStreams`；consumer 已运行 | VERIFIED（可运行 TFM） |
| `NpoiRelationBinder.Bind` | `Bing.Offices.Tests.ExcelWorkbookRequestTest.Import_RelationDelegateFailure_ShouldPreserveOriginalExceptionType`（五类委托点） | Integration XLS/XLSX 真实导入 | VERIFIED（net6/net8 Unit + Integration） |
| CSV 公式注入 | `CsvTest.EntityPipeline_FormulaPrefixesAfterWhitespace_ShouldEscapeWithoutChangingNegativeNumbers`、`EntityPipeline_PreserveFormulaPolicy_ShouldKeepOriginalText`、RFC4180/culture 用例 | Integration `AddBingOfficesNpoi_CsvServices_ShouldRoundTripStream`；Docs CSV fence | VERIFIED（可运行 TFM） |
| `CsvImportOptions.Validate` | `CsvTest.EntityPipeline_InvalidUniqueOptions_ShouldThrowAtValidationBoundary` | CSV integration | VERIFIED |
| 资源限制/ResourceProbe | `ExcelP0RegressionTest.Import_ResourceProbe_ShouldRunInIndependentProcess`；`artifacts/excel-resource-probe-rerun.jsonl`；`artifacts/resource-probe-rerun.jsonl` | Excel 七模式独立 child process；mapping/unique 为另一套 16 场景 child workload；当前未覆盖任意真实 DOM 输入上限 | PARTIAL / Excel 7/7 + mapping/unique 16/16 |
| Failure Workbook/AtomicFileCommitter | `ExcelP0RegressionTest` Failure Workbook 相关 14 个专项测试；Integration 锁文件/复制失败 | Windows 临时目录、锁文件、取消和目标保护 | VERIFIED（专项 14/14；Integration net6/net8） |
| API 删除/分类 | `PublicApiContractTest.PublicApi_ReleaseAssemblies_ShouldMatchApprovedBaseline`、`PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`、exact member test | isolated nupkg consumer 无 ProjectReference | VERIFIED（正式 baseline 已批准；API compare 退出码 0；可运行 TFM Unit 全绿） |
| DI 注册 | `StreamPipelineTest.AddBingOfficesNpoi_*`；Integration/Docs/consumer | replacement、链式返回、服务解析和 package-only 调用 | VERIFIED |
| `MaxSerializedBytes` | `ExcelP0RegressionTest.Import_FailureWorkbook_ExceedingMaxSerializedBytes_ShouldNotWriteDestination`、temporary cleanup/copy failure | `artifacts/package-consumer-rerun2` package-only consumer 构造并校验新属性 | VERIFIED（net8专项 14/14；consumer 短路径 restore/build/run 通过，长路径缓存 MSB3106） |
| Benchmark | `StreamPipelineBenchmarks` 9/9 ShortRun；补充 Mapping/Failure/Dynamic/Unique/Tenant artifacts | ResourceProbe/尾延迟独立 JSONL | PARTIAL / `UNAPPROVED`；未形成批准预算门禁 |

## Round 5 回归绑定

| 验证批次 | 真实结果 | 证据 |
| --- | --- | --- |
| Unit net8 | `382 total / 381 passed / 1 failed`；唯一失败为正式 API hash mismatch | `tests/Bing.Offices.Tests/TestResults/review-fix-round5-unit-net8.trx` |
| Unit net6 | `382 total / 381 passed / 1 failed`；唯一失败为正式 API hash mismatch | `tests/Bing.Offices.Tests/TestResults/review-fix-round5-unit-net6.trx` |
| Integration net8/net6 | `15/15 + 15/15 = 30/30` | `integration-test-report.md` 与 Round 5 TRX |
| Docs net8 | `11/11` | `docs-test-report.md` 与 Round 5 TRX |
| 显式 CSV/校验专项 | `6/6` | `tests/Bing.Offices.Tests/TestResults/review-fix-round5-focused-csv-net8.trx` |
| package-only consumer | Round 5 nupkg，短路径 restore/build/run `0/0/0`，输出 `package-consumer-ok` | `package-consumer-report.md`、`artifacts/packages-round5` |

## Round 6 回归绑定

| 验证批次 | 真实结果 | 证据/限制 |
| --- | --- | --- |
| Release solution build | `dotnet build Bing.Offices.sln -c Release --no-restore` 成功，`0 error / 28 warning` | 包含 Benchmark、ResourceProbe；warning 主要为 netcoreapp3.1 依赖支持、nullable、obsolete、analyzer |
| StreamPipeline Unit net8/net6 | 两个 TFM 均 `90/90` | 覆盖 NPOI import/export、动态列、表头、取消、资源、Failure Workbook 相关管线；完整 Unit 仍受 formal API hash 阻断 |
| Integration net8/net6 | `15/15 + 15/15` | 两个可运行 TFM 均通过 |
| Docs consumer net8 | `11/11` | 动态列 Markdown 已改为公开 `ICsvImporter` + DI，所有围栏通过 |
| API contract targeted | 类型清单、分类、NPOI public exact、结构异常回归通过；formal member hash 仍失败 | `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot` 的 Abstractions actual=`7F9A2AA819E94B3838097DF2FF374A934CF7F35F3D2E91F3D1DB790F22972943`，expected 旧 baseline；未更新 hash |

## 正式 API baseline 批次

| 验证批次 | 真实结果 | 证据 |
| --- | --- | --- |
| API snapshot compare | `netcoreapp3.1`、`net6.0`、`net8.0` 全部通过，退出码 `0` | `artifacts/api-snapshot-formal-baseline-20260903.json`、`artifacts/api-snapshot-formal-20260903/api-snapshot-*.json` |
| Unit net8 | `384 total / 384 passed / 0 failed` | `tests/Bing.Offices.Tests/TestResults/api-baseline-net8-final-rerun.trx` |
| Unit net6 | `384 total / 384 passed / 0 failed` | `tests/Bing.Offices.Tests/TestResults/api-baseline-net6-final-rerun.trx` |

正式 API hash：Abstractions `7F9A2AA819E94B3838097DF2FF374A934CF7F35F3D2E91F3D1DB790F22972943`；Core `B3661970BBE5AECC06DAD57B1E3F960FA77E70C4D2E66B2DA4910F7823AA2BB6`；NPOI `DA163263804A964D8AC2A13D78D6B3858256171CE7729841690FDB56F602CEEE`。
