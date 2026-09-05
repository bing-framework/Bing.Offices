# 最终发布门禁报告

## 结论

`BLOCKED / No-Go`。

本任务已完成可安全执行的核心实现、回归验证、独立 Review、Review 修复和证据整理，但未满足发布所需的全部门禁。禁止将当前工作树视为可发布包，也未执行 publish、tag、commit、push 或 PR。

## 已通过项目

- 统一异常合同已接入 Core/CSV/NPOI/Failure Workbook/Atomic File Commit 主链；保留 inner exception，取消、参数和致命异常边界已覆盖。
- 日期解析已统一为 Core parser：默认 ISO `yyyy-MM-dd`、InvariantCulture、`DateTimeKind.Unspecified`、DateTimeOffset 显式或固定 offset、1900/1904 serial 和 Workbook DATE/TIME façade。
- XLSX ZIP preflight 已在 NPOI `WorkbookFactory.Create` 前执行；entry、解压大小、压缩比、关键 XML 部件、重复 entry、路径和 DTD/entity 检查已实现。
- NPOI 六类 Provider User API 已对外可见，内部辅助扩展未泄漏；package-only consumer net6/net8 均成功。
- 生产程序集之间无 `InternalsVisibleTo`，仅保留测试友元。
- Review Round 1 的 `FIX-001` 已修复：`MaxZipCompressionRatio = null` 现在表示关闭该限制并通过配置验证；`0`、负数、NaN、正/负无穷仍被拒绝。Round 2 独立复核为 `PASS_WITH_ISSUES`，无新的 P0/P1 代码问题。

## 验证结果

| 门禁 | 结果 | 证据 |
| --- | --- | --- |
| Release solution build | PASS，0 error，22 warnings | `dotnet build .\\Bing.Offices.sln -c Release --no-restore` |
| Unit net8 | `414/415`，唯一失败 formal API hash | `unit-test-report.md`、`rc-hardening-review-fix-unit-net8-final.trx` |
| Unit net6 | `414/415`，唯一失败 formal API hash | `unit-test-report.md`、`rc-hardening-review-fix-unit-net6-final.trx` |
| Unit netcoreapp3.1 | `414/415`，唯一失败 formal API hash | `unit-test-report.md`、`rc-hardening-review-fix-unit-netcoreapp31-final.trx` |
| ZIP/资源 Unit | `16/16`，net8/net6/netcoreapp3.1 | Review fix TRX |
| Integration net8 | `15/15` | `rc-hardening-review-fix-integration-net8-final.trx` |
| Integration net6 | `15/15` | `rc-hardening-review-fix-integration-net6-final.trx` |
| Docs net8 | `11/11` | `rc-hardening-review-fix-docs-final.trx` |
| Excel ResourceProbe | 12 child scenarios；5 个 Preflight reject | `artifacts/excel-resource-probe-rerun.jsonl` |
| Mapping/Unique ResourceProbe | `16/16` | `artifacts/mapping-resource-probe.jsonl` |
| MappingValidation Benchmark | `10/10` ShortRun | `benchmark-report.md`、BDN results/log |
| Package-only consumer | net6/net8 restore/build/run PASS，输出 `package-consumer-ok` | `package-consumer-report.md` |
| `get_errors` | No errors found | 源码/测试目录诊断 |
| `git diff --check` | PASS；无 whitespace error，仅有 CRLF/LF 转换提示 | 收口阶段执行 |

## 阻断项

1. **Formal API approval 未完成。** candidate 与 formal baseline hash 不一致：Abstractions `F66D...` vs `7F9A...`，Core `AE89...` vs `B366...`，NPOI `1C5E...` vs `DA16...`。本任务未修改 formal baseline，也未修改测试 hash 以掩盖差异。
2. **Review 需要最终 sign-off。** Round 1 的 MUST_FIX 已完成，Round 2 独立复核为 `PASS_WITH_ISSUES`；发布级 API/预算/矩阵问题仍需维护者签核。
3. **性能/资源预算未批准。** 当前 Benchmark 和 ResourceProbe 是有效证据，不是经批准的吞吐、分配、LOH、工作集或尾延迟门禁。
4. **资源矩阵不完整。** Failure Workbook 双 DOM、完整 NPOI DOM、100k/1M Excel/CSV、取消延迟及 XLS/OLE 内部结构保护未形成完整当前任务矩阵。
5. **当前任务 Benchmark 不完整。** 仅 MappingValidation `10/10` ShortRun；历史 StreamPipeline 100k 的 3 个异常尚未形成可采信的当前任务解释，历史 Failure Workbook 约 2.5 GB 分配不能代替当前测量。
6. **兼容性治理未完全收口。** DataTable 显式兼容类、Office exception 家族、UniqueTracker 和剩余 public execution detail 仍需逐符号批准或迁移；不把 `EditorBrowsable(Never)` 当作 internal。
7. **交付环境证据有限。** package-only 已在隔离缓存 net6/net8 成功，但同版本本地 nupkg 重发不是 feed immutability、clean clone 或正式发布证明；历史深路径曾有 `MSB3106`。
8. **日期/文档残余风险。** 导出 Workbook 元数据仍使用 `DateTime.Now`；不影响当前输入日期 parser，但不满足更强的导出元数据完全确定性要求。DateOnly 保持不支持，最终用户文档需随 API approval 再确认。

## 需要维护者动作

- 审批或拒绝 `api-diff.md` 中的成员级 breaking diff，并在批准后独立更新 formal baseline、重跑 API compare 和全量 Unit。
- 对性能与资源预算给出明确数值和 workload 范围，或记录正式 waiver。
- 完成独立 Review 最终签核；若产生新的 MUST_FIX/SHOULD_FIX，重新进入对应 Review 修复轮次。
- 决定 DataTable/Office exceptions/public execution detail 的兼容策略和后续 task 归属。
- 在 clean clone 或正式 feed staging 环境复核包资产、TFM 和 lockfile。

## 文档索引

- `api-diff.md`
- `deprecated-removal.md`
- `unit-test-report.md`
- `integration-test-report.md`
- `docs-test-report.md`
- `package-consumer-report.md`
- `benchmark-report.md`
- `resource-report.md`
- `review.md`
- `review-round2.md`
- `progress.md`
- `execution.md`

## Git 与发布安全

- 未执行 `git add`、`git commit`、`git push`。
- 未创建 tag、PR 或执行 publish。
- 未执行 `git reset`、`git clean` 或覆盖未知修改。
