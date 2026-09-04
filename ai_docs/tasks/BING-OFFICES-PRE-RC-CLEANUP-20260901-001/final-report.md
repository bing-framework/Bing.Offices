# 最终执行报告

## 任务与结论

- Task ID：`BING-OFFICES-PRE-RC-CLEANUP-20260901-001`
- 执行结论：`BLOCKED`
- 发布判定：`No-Go`
- 执行时间：2026-09-01
- 执行器：Copilot plan-executor
- 计划：`plan.md`
- 独立 Review：`review.md`
- 发布清单：`release-checklist.md`

本轮持续执行已批准计划，并完成用户批准的正式 API baseline 收口。API compare、net8/net6 Unit 和最终 Release build 均已通过；整体仍因缺失 TFM runtime、完整性能/资源门禁、长路径 package consumer、剩余 API 治理和 clean-clone 交付证据保持 `BLOCKED/No-Go`。`review.md` 未修改。

## 正式 API baseline 收口批次（2026-09-03）

- 用户批准将当前 Release API 纳入正式 baseline，批准范围限于已记录的 Breaking Change 和 API 收敛。
- 正式 baseline：`artifacts/api-snapshot-formal-baseline-20260903.json`。
- 当前快照：`artifacts/api-snapshot-formal-20260903/api-snapshot-*.json`。
- API compare 退出码：`0`；`netcoreapp3.1`、`net6.0`、`net8.0` 全部通过。
- 正式 hash：Abstractions `7F9A2AA819E94B3838097DF2FF374A934CF7F35F3D2E91F3D1DB790F22972943`（723 members）；Core `B3661970BBE5AECC06DAD57B1E3F960FA77E70C4D2E66B2DA4910F7823AA2BB6`（194 members）；NPOI `DA163263804A964D8AC2A13D78D6B3858256171CE7729841690FDB56F602CEEE`（1 member）。
- 与 Round 5 快照的成员差异仅为已批准 Settings 删除、Core execution detail internal 化及 Provider `cacheCapacity` 尾参数收敛；未发现未批准的公共 API 新增。
- 最终 Release build：`dotnet build .\Bing.Offices.sln -c Release --no-restore`，退出码 `0`，`0 error / 15 warning`。
- Unit net8：`384/384`，TRX：`tests/Bing.Offices.Tests/TestResults/api-baseline-net8-final-rerun.trx`。
- Unit net6：`384/384`，TRX：`tests/Bing.Offices.Tests/TestResults/api-baseline-net6-final-rerun.trx`。
- 本批次关闭 API baseline 和可运行 TFM Unit 的 hash 阻断，不改变整体 RC 判定。

## 已完成事项

### 生产与测试修复

- Mapping Document request configuration 按方向合并并进入最终 Plan；
- v1 JSON/XML 迁移的非目标方向保持 `null`，并保留 XML DTD/外部实体防护和流所有权合同；
- Relation Binder 解包 `TargetInvocationException` 并重抛原异常；
- CSV 公式防护覆盖 BOM、控制字符、ASCII/Unicode whitespace，保留 `Escape`/`None` 语义和负数行为；
- `CsvImportOptions` 增加 unique limit/comparison 边界校验；
- ResourceProbe 改为独立 child process，不预先打开同一 Workbook；
- `AddNpoi` 删除，唯一 DI 入口为链式 `AddBingOfficesNpoi`；
- 未知 Workbook/Sheet 能力改为 `NotSupportedException`；
- `MaxBytes` 更名为 `MaxSerializedBytes`，并完成 Failure Workbook 输出限制、取消、临时文件、清理诊断和目标流保护测试；
- NPOI 2.7.4 capability 审计确认 HSSF `Hidden`/`Collapsed` 不支持而 `ZeroHeight` 支持，XSSF 三项支持，三处 narrow fallback 保留。

### 可复核验证

- Release solution build：成功；
- Integration net6/net8：`30/30`；
- Docs net8：`11/11`；
- Failure Workbook 专项：`14/14`；
- DI 定向测试：`6/6`；
- Excel ResourceProbe：`7/7` child scenarios；另有 `16/16` mapping/unique benchmark child scenarios，两套 workload 已分离记录；
- StreamPipeline BenchmarkDotNet：`9/9` ShortRun；
- package-only consumer：`artifacts/package-consumer-rerun2` 无 `ProjectReference`，使用 Round 5 `packages-round5` 的本地 2.0.0 nupkg 和 `C:\nupkg-cache-round5` 时 restore/build/run 均退出码 `0`，输出 `package-consumer-ok`。

## 部分与阻断事项

1. **TFM runtime 缺失。** netcoreapp3.1、net5.0、net7.0 未安装 runtime，相关 Unit 未运行，不声明通过。
2. **性能门禁未完成。** ShortRun 仅为证据；完整 Benchmark 矩阵未形成，尾延迟 budget 为 `UNAPPROVED`，100,000 行 Import 记录 3 个异常需复核。
3. **ResourceProbe 边界不足。** 当前已取得七个真实 Excel child-process 样本，并另行保留 mapping-plan/unique-tracker 子进程样本；仍不能证明任意 XLS/XLSX、压缩输入、NPOI DOM 或 Failure Workbook 双 DOM 的硬内存上限。
4. **Package consumer 环境限制。** 短路径缓存成功；任务深路径缓存 restore 后 build 触发 SDK/MSBuild `MSB3106`，随后产生 DI Abstractions 解析错误。该环境失败已如实保留，不能写成普通 consumer 通过。
5. **API 分层/删除候选未最终收口。** 已批准的 legacy converter、旧 validation attributes、CSV 全局状态/旧隐式重载、Office exceptions、Settings 及部分 Core execution detail 已收敛；DataTable 显式兼容类、UniqueTracker 和剩余 public execution detail 仍需逐符号治理与迁移闭环。
6. **P3 重构/ADR 未完成。** Import/CSV/Mapping 拆分和完整异步/所有权 ADR 未全部完成，保持 TODO/partial。

## 证据索引

- 基线：[baseline.md](baseline.md)
- 进度：[progress.md](progress.md)
- API 差异：[api-diff.md](api-diff.md)
- 删除与兼容：[deprecated-removal.md](deprecated-removal.md)
- 测试矩阵：[test-matrix.md](test-matrix.md)
- Unit：[unit-test-report.md](unit-test-report.md)
- Integration：[integration-test-report.md](integration-test-report.md)
- Docs：[docs-test-report.md](docs-test-report.md)
- Package consumer：[package-consumer-report.md](package-consumer-report.md)
- Benchmark 计划：[benchmark-plan.md](benchmark-plan.md)
- Benchmark 报告：[benchmark-report.md](benchmark-report.md)
- 资源报告：[resource-report.md](resource-report.md)
- 独立 Review：[review.md](review.md)
- 发布清单：[release-checklist.md](release-checklist.md)
- 原始 TRX/JSONL/BDN：任务 `artifacts/` 目录及仓库 `BenchmarkDotNet.Artifacts/results/`

## 计划偏差

- 计划中的 P3 大规模职责拆分未执行：当前 P0/P1 修复、测试和发布证据优先，避免在 API/性能门禁未收口时扩大行为风险。
- API formal baseline 已按用户批准更新并完成 compare/Unit 验证；剩余 API 分层治理仍独立保持阻断。
- package consumer 采用短路径 NuGet cache 作为最小可逆环境隔离；同时保留深路径 `MSB3106` 失败证据，不将 workaround 隐藏为无条件通过。

## Build / Test / Format

- Release build：PASS；
- Unit net8：`384 total / 384 passed / 0 failed`；Unit net6：`384 total / 384 passed / 0 failed`；
- Integration：PASS，net6/net8 合计 `30/30`；
- Docs：PASS，net8 `11/11`；
- `get_errors`：源码、测试和构建相关诊断为 0；
- `git diff --check`：通过，仅有 CRLF/LF 提示；
- Lint：项目未发现独立 lint 入口，未伪造 lint 通过。

- Excel ResourceProbe：`7/7` child scenarios，非完整 DOM 安全证明；mapping/unique benchmark ResourceProbe：`16/16`，不可替代 Excel workload；

## API / 配置变化

- `ExcelImportFailureOptions.MaxSerializedBytes` 替代 `MaxBytes`；语义仅为 Failure Workbook 序列化输出限制；
- `ExcelNpoiServiceCollectionExtensions.AddBingOfficesNpoi(IServiceCollection)` 替代旧 `AddNpoi`，返回同一 services；
- 其它候选兼容 API 未机械删除，等待逐符号治理和维护者批准。

## 安全与风险

- 未发现本轮新增 secret、网络依赖、catch-all 吞异常、外部实体解析放开或不受控目标写入；
- CSV 注入防护不宣称覆盖所有表格执行风险；
- `MaxInputBytes`、NPOI DOM、解压内容、实体对象图、Failure Workbook 序列化输出和进程工作集保持分离表述；
- 直接目标 Stream 失败可能部分写入，File API 继续使用临时文件和原子提交；
- 旧 `artifacts/package-consumer` 失败后旧二进制输出不作为成功证据。

## Reviewer 注意事项

- 复核正式 API baseline 是否与当前 Release DLL、快照和成员差异绑定；
- 确认当前 Unit/API compare 通过来自用户批准的 Breaking Change 和真实 compare，不是忽略或条件跳过；
- 检查 Benchmark 报告是否保留 100k Import 的 3 个异常和 `UNAPPROVED` 状态；
- 检查 ResourceProbe 是否被误读为任意 Workbook 内存硬上限；
- 检查 package consumer 必须使用最新 nupkg、无 ProjectReference，并区分短路径成功与长路径 `MSB3106`；
- 复核 public execution detail/legacy compatibility 的删除或保留批准记录。

## 建议提交分组（未执行）

1. `fix(core): correct mapping migration relation and csv safety contracts`
2. `fix(npoi): harden failure workbook and fluent registration API`
3. `test: add regression integration docs and package consumer coverage`
4. `docs: record release candidate evidence and remaining gates`

以上仅为建议，不代表已执行 Git 操作。

## Git 状态与安全边界

- 未自动执行 `git add`；
- 未自动执行 `git commit`；
- 未自动执行 `git push`；
- 未自动创建 PR、tag 或发布；
- 未执行 `git reset`、`git clean` 或覆盖未知用户修改；
- 当前工作树包含本任务生产/测试/文档修改和任务证据目录，正式交付前需由授权人员审阅、staging 和提交。

## 最终结论

`BLOCKED` / `No-Go`。当前环境可安全完成的实施、验证、正式 API baseline 收口和证据整理已完成；可运行 TFM 的 API compare/Unit 门禁已通过，但缺失 runtime、正式性能预算、资源覆盖、剩余 API 治理和交付复现条件仍未满足，不能进入发布。

## Round 5 补充

Round 5 在用户明确授权清理已迁移的 `[Obsolete]` API 后完成以下 breaking cleanup：

- 删除 `ICellValueConverter` 及其 NPOI importer/materializer legacy bridge 和测试专用实现；
- 删除 `RequiredAttribute`、`RegexAttribute`、`RangeAttribute`、`MaxLengthAttribute`、`DateTimeAttribute`、`DuplicationAttribute` 六个旧 validation attributes，并将生产规则、测试模型、Docs/API contract 和 package consumer 迁移到 `Excel*` 替代 API；
- 删除 `CsvSeparatorCharacter`、`CsvQuoteCharacter` 及依赖全局状态的旧隐式 DataTable 重载，保留显式 delimiter/quote API。

Round 5 重新生成 `artifacts/api-snapshot-review-fix-round5`：Abstractions 新 hash 为 `C176D71B0025C1F28F010BF05667588898A4D0EA4F847CD65658D9737D800313`，Core 新 hash 为 `410F6A0F6CF64B41C3AB141AECFB2E1606C9B116EF2BD56C9A868FE02DC8FB68`；正式 baseline 未更新。net6/net8 Unit 均为 `382 total / 381 passed / 1 failed`，唯一失败为正式 API hash mismatch；Integration `30/30`、Docs `11/11`、显式 CSV/校验专项 `6/6`、Round 5 package-only consumer `package-consumer-ok`。

Round 5 不扩展删除 Office exceptions、Settings、UniqueTracker、public execution detail 或 P3 重构。`decisions.md` D-013 已逐项记录当前无具名批准人、无后续 taskId、无迁移期限、无 waiver 的真实状态及 RC `No-Go` 影响，因此这些项目仍为 `BLOCKED/PARTIAL/TODO`，不能写成已批准延期。独立 `review.md` 保持原样，等待再次验收。
