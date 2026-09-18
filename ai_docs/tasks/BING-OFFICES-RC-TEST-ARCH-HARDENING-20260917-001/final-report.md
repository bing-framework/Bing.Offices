# 最终报告

## 代码整改

- 状态：代码整改 PARTIAL，正式发布 PARTIAL（职责隔离、六项目双 TFM、Provider 集成回归、精确 IVT、CI Large/benchmark 入口和 provider-shared preflight 已完成；未经可信 baseline 支持的 RawDate 过滤/重复 Query、Relation 缓存优化及 Importer 物理列预绑定均已撤回；性能收益未宣称）。
- Common 只引用 Abstractions/Core；Provider Unit/Integration 和 Aggregate 依赖符合矩阵。
- 默认 ValidateMode、ExcelImportValidationMode、先注册优先 DI、公共 API/TFM/版本文件保持不变。
- Round 3 六职责历史 TRX：Common Unit 196、NPOI Unit net8 533、MiniExcel Unit 36、Aggregate Integration 29、NPOI Integration 28、MiniExcel Integration 7；NPOI Unit net6 历史全量为 532/533（1 个中途取消失败）。Round 7 NPOI Unit 双 TFM 各 539/539（含三组 NPOI 分支回归），MiniExcel Unit 双 TFM 各 39/39。NPOI/MiniExcel 各另有 500K/1M `Category=Large` Case，正式 workflow/生产资源证据仍未验证。
- `git diff --check`：PASS。

## Round 4/5 修复复验

- MiniExcel Unit Round 5 双 TFM：39/39 通过；新增 Relation 重复 ParentKey 委托调用/异常边界回归，并保持空 child 与首匹配短路回归。
- 方法身份对账使用同一 UTF-8 锚定解析器：HEAD 569、当前 592、Added 23、Removed 0；完整去向见 `method-identity-audit.md`，不与 TRX 展开 Case 混用。
- Hotspot 历史样本保留：`artifacts/benchmarks/review-fix-round4-hotspot-after2.jsonl`。RawDate 过滤/重复 Query 与 Relation 缓存/索引均已撤回；RawDate before 是调查性 replay，历史 after 不再作为当前候选收益证据。

## 发布状态

- 状态：PARTIAL。
- API snapshot compare：NOT_VERIFIED（candidate identity mismatch）。
- 资源批准、500K/1M、真实 2 CPU/4 GiB、外部 CI：NOT_VERIFIED/BLOCKED。
- RawDate/Relation hotspot：Round 3/4 文件保留为历史调查资料，不作为当前候选收益证据；当前探针仅保留可运行的全量 numeric-cell 调查 workload；Importer 列预绑定已撤回，绝对物理列错误定位由直接回归覆盖。
- Provider comparison 仍只有当前候选 after 样本，不能计算 Provider 间收益；旧 after artifact 保持原样，新采样使用独立进程、候选 identity 和真实 `PeakWorkingSet64`。

## 证据索引

- `baseline.md`
- `test-migration-matrix.md`
- `provider-test-matrix.md`
- `unit-test-report.md`、三个 integration reports
- `api-diff.md`
- `package-consumer-report.md`
- `benchmark-report.md`
- `resource-report.md`
- `symbol-test-map.md`
- `method-identity-audit.py`、`method-identity-audit.md`
- `review.md`
- `execution.md`
