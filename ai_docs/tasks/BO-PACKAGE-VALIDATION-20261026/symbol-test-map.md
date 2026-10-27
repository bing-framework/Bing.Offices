# 能力与验证追溯

本轮不新增或删除生产 API。下表将既有生产入口及新增验证工具映射到独立包消费断言；所有消费者使用 PackageReference，在仓库外的干净目录执行。

| 最终符号 / 入口 | 关键行为 | 消费项目 / 测试方法 |
| --- | --- | --- |
| Core 服务扩展 AddBingOfficesCore | 默认 CSV 注册可解析；不注册 Excel Provider | Bing.Offices.Consumer.Core Program，服务解析断言 |
| ICsvExporter.Export / ExportAsync（经 ExportToBytes / ExportToBytesAsync） | 完整 UTF-8 CSV 输出一致 | Bing.Offices.Consumer.Core Program，expectedCsv 与字节序列断言 |
| ICsvImporter.Import / ImportAsync（经字节扩展） | 完整实体、普通行校验错误后继续；统一结果摘要 | Bing.Offices.Consumer.Core Program.AssertSuccess / AssertValidationFailure |
| CsvImportResult.IsSuccess / IsComplete / StopReason | 无错误成功；普通校验错误完成但不成功 | 同上，包含完整 Validation 错误字段 |
| ExcelSingleSheetImportExtensions 六个 ImportSheet 入口 | XLS/XLSX × Stream/Bytes/File × Sync/Async，成功与结构化表头失败 | Bing.Offices.Consumer.Net6/Net8 调用 ExcelSingleSheetPackageChecks.RunAsync / InvokeAsync |
| ExcelSingleSheetImportResult.Items / SourceRows / Errors / IsSuccess / IsComplete / StopReason | 完整实体和源行，完整 InvalidHeader 错误及截断状态 | ExcelSingleSheetPackageChecks.AssertSuccess / AssertInvalidHeader |
| ApiSnapshot CLI --candidate | 身份来自候选，成员/审批来自批准基线 | Verify-PackageCandidate.ps1：capture / compare / reject-api-mismatch |
| CandidateIdentityVerifier.Validate | 缺失清单拒绝；程序集身份篡改拒绝；保留源、包和审批检查 | reject-missing-candidate / reject-bad-identity；CandidateIdentityContractTest.Run |
| Verify-PackageCandidate.ps1 / Invoke-DotNet | 子命令非零即失败，反例要求特定诊断；独立包源和缓存 | 完整四目标消费链路；三项 CLI 拒绝用例 |
| Verify-PackageCandidate.ps1 恢复图检查 | 不含 ProjectReference；Core 无 NPOI；Office 包实际来自本轮源 | 每个消费者 project.assets.json 与 .nupkg.metadata 的运行时断言，summary.restoredSources |
| .github/workflows/ci.yml Verify package candidate | 调用统一脚本并归档成功/失败证据 | 本地同命令验证 + YAML 静态检查；远端 run 另行验收 |

既有 Workbook、直接 NPOI Provider 和 NPOI 扩展消费验证保留在 Net6/Net8 的 Program.cs 中。之前生产职责测试及成员快照映射继续见 BO-API-USABILITY-20261025/symbol-test-map.md。
