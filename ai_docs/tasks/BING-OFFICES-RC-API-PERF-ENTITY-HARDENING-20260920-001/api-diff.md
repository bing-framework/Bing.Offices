# API 差异

来源：`artifacts/api/compare/current-20260921-candidate/api-diff.json`，候选程序集来自当前 `output/release`，双 TFM 结果相同；历史 compare 产物不作为当前审批输入。

## Breaking Diff

新增：

- `Bing.Offices.Imports.ExcelValidationFailureMode`（`Continue`、`StopOnFirstFailure`）
- `ExcelSheetImportBuilder<TItem>.Validate(ExcelValidationFailureMode)`
- `ExcelSheetImportRequest.ValidationFailureMode`
- `Bing.Offices.Entities.ExcelEntity*` immutable layout/result/template contracts。
- `Bing.Offices.Extensions.ExcelEntityExtensions` 的 Entity/Template 同步异步入口。
- `Bing.Offices.Entities.IExcelEntityImporter/IExcelEntityExporter` 和 `Bing.Offices.Providers.ExcelProviderCapabilities/IExcelProviderCapabilities` Provider SPI。
- NPOI/MiniExcel provider 的 capability 声明；MiniExcel 对未支持的 Entity/Template 能力 fail-fast。

删除：

- `Bing.Offices.Imports.ValidateMode` 及其字段
- `ExcelSheetImportBuilder<TItem>.Validate(ValidateMode)`
- `ExcelSheetImportRequest.ValidateMode`

旧基线的 Abstractions 成员数量为 `786`；当前候选包含上述新增契约以及 rename，因此 snapshot 差异不是运行时行为错误。现有 `build/api-snapshot-baseline.json` 仍是旧候选，未获得本任务成员级批准，因此没有覆盖该 baseline。`PublicApiContractTest` 在 net6/net8 各报告 `27 passed / 2 failed`，失败均为 baseline/snapshot mismatch。

## Approval Gate

现有 `ai_docs/tasks/BO-RC-20260908-002/api-breaking-approval.md` 只批准四个 CSV/Excel 文件导出扩展删除，不覆盖本次 `ValidateMode` 重命名。故 API snapshot gate 状态为 `NOT_VERIFIED`，需要成员级审批后重新生成 baseline 并重跑测试。

2026-09-21 复核：旧 `output/release` capture 与候选 manifest 的 source hash 不一致问题已通过当前工作树重新 build/pack/capture 形成独立刷新记录。当前 capture 位于 `artifacts/api/current-20260921-candidate`，source-scope hash 为 `E4447DEB168499218EE856357197D80A96208F625F79F4DCABDCB1E4AB101346`，并与 `artifacts/packages/candidate-20260921` 的四个包身份一致；明细见 `artifacts/candidates/.../manifest-current-20260921.md`。旧 baseline 仍会因未批准的成员差异而拒绝 compare，未修改 `build/api-snapshot-baseline.json`；benchmark/resource 旧产物尚未重新封存到该刷新候选。

当前刷新候选对旧 baseline 的 compare 仍按预期失败：失败项包括 baseline candidate identity/hash 不一致，以及 `ValidateMode` rename、Entity/Template/SPI 新增成员造成的 snapshot diff。该失败是 API approval gate 的证据，不是当前 build/package/API capture 内部身份不一致；baseline 仍保持原文件和原哈希。
