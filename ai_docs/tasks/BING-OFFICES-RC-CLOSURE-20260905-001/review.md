# 独立代码复审报告

<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BING-OFFICES-RC-CLOSURE-20260905-001
AI_REVIEWED_AT: 2026-09-06T13:09:05+08:00

## 复审结论

- 结论：`NEEDS_FIX` / `No-Go`。
- 开放 P0：0。
- 开放 P1：2（FIX-002、FIX-005）。
- 开放 P2：2（FIX-006、FIX-008）。
- 已关闭：FIX-001、FIX-003、FIX-004、FIX-007。
- 本次复审未修改生产代码，未 commit、push、创建 PR、tag 或 publish。

## P0 Findings

无开放 P0。

## P1 Findings

### FIX-002：正式 Public API 门禁仍失败，候选快照不足以独立承担 APICompat

- 状态：`OPEN`。
- 位置：`tests/Bing.Offices.Tests/PublicApiContractTest.cs:635-677`、`build/ApiSnapshot/PublicApiSnapshot.cs:58-100`、`api-diff.md`。
- 证据：最终冻结后的三个 Unit TFM 均为 471 total / 470 passed / 1 failed，唯一失败仍是 Abstractions formal hash：expected `7F9A...`，actual `545834...`。仓库没有历史成员 baseline JSON，无法进行成员级批准。
- 工具缺口：当前 canonicalizer 不编码 abstract/sealed、泛型约束、参数名/默认值、属性访问器可见性和关键 attribute；本轮 `BingOfficesException` 改为 abstract 就不会改变该 hash。
- 影响：Breaking Change 尚未完成可追溯审批，Unit 报告仍为 FAILED，最终 APICompat 门禁不满足。
- 解除条件：恢复正式 baseline 或取得维护者成员级审批；使用覆盖上述兼容维度的 APICompat/快照后更新批准基线，并重跑三个 TFM。

### FIX-005：Benchmark/Resource 已扩充真实矩阵，但仍无可比 baseline、批准预算和完整资源证据

- 状态：`OPEN`，已有实质进展。
- 位置：`benchmark-report.md`、`resource-report.md` 及 `artifacts/benchmark/*-shortrun`。
- 已验证：新增 Excel、CSV、Failure Workbook 1k/10k/100k 公开调用链 ShortRun；原始 CSV/JSON/Markdown artifacts 存在，报告数字与原始 CSV 抽查一致。报告如实披露高分配与 Gen2，没有宣称低 GC。
- 剩余缺口：没有历史可比 baseline 和维护者批准预算；ShortRun 的 Error 很大，部分 Import 有内部 exception 事件；缺少图片/模板/样式/验证完整矩阵、Failure Workbook 双 DOM 的独立 PeakWorkingSet/LOH 保留量，以及取消延迟分位数。
- 影响：无法判定性能回退和资源是否可接受，最终门禁第 10、11 项不满足。
- 解除条件：维护者给出预算与批准人；在固定环境运行可比 baseline、完整 Benchmark/Resource 矩阵并审查高分配、Gen2 和内部 exception 事件。

## P2 Findings

### FIX-006：职责拆分已有独立落点，但大类拆分仍未完成

- 状态：`PARTIALLY_FIXED`，由 P1 降为 P2。
- 已完成：新增 internal `NpoiFailureWorkbookDiagnostics`，从 writer 中抽离 Provider 行元数据降级、sink 隔离和结构化 Trace；三条职责级测试通过。
- 剩余：`NpoiFailureWorkbookWriter` 仍约 857 行，`CsvEntityPipeline` 863 行，`NpoiExcelImporter` 630 行，`NpoiExcelExporter` 565 行；计划中的 failure workbook 构建/复制/摘要/序列化/临时文件以及 CSV Import/Export/Conversion/Validation 分离未全部落地。
- 判断：当前有定向和全量回归支持，未发现由此产生的发布正确性缺陷，故降为维护性 P2；不能在执行报告中把 Phase 3 全部职责拆分写成 VERIFIED。

### FIX-008：TryAddPicture 缺少可恢复失败后副作用的合同测试

- 状态：`OPEN`。
- 位置：`SheetExtensions.Picture.cs` 的 `TryAddPicture(byte[])`。
- 证据：方法先执行 `Workbook.AddPicture`，之后 `CreatePicture/Resize` 的 `ArgumentException` 或 `InvalidOperationException` 会返回 false。现有 Unit/PackageConsumer 覆盖预校验无修改和成功路径，没有覆盖后半段失败时是否残留 picture 数据或 drawing。
- 影响：调用方看到 false 时的部分修改语义仍未冻结。
- 解除条件：增加可控失败注入测试；明确允许部分修改还是回滚，并同步文档。

## 已关闭 Findings

### FIX-001：CLOSED

- `NpoiFailureWorkbookDiagnostics` 在没有 sink 时写入包含 code/property/row/exception 的结构化 Trace；sink 的普通异常同样写独立 Trace，取消和 OOM 保持原样。
- 独立复跑三条职责测试 3/3 通过；最后拆分后的全量回归也完成。

### FIX-003：CLOSED

- 新增 XLS/XLSX 导入→导出→关闭重开→再导入测试，断言字符串单元格、值和 offset。
- 独立在 `TZ=UTC` 与 `TZ=Pacific Standard Time` 复跑 DateTimeOffset 专项，各 6/6；文档描述与实际矩阵一致。

### FIX-004：CLOSED

- 最终冻结包目录为 `artifacts/packages-rc-final`，消费者 assets 指向全新 `.packages-rc-final` 缓存且无 Bing ProjectReference。
- 独立计算三个 nupkg SHA-256 与报告一致：Abstractions `5DF62D...`、Core `6CDE11...`、Npoi `8FA169...`。
- 包内五个 DLL 与当前 `output/release` 全部 hash match；独立运行 netcoreapp3.1、net6.0、net8.0 均输出 `package-consumer-ok`。
- PackageConsumer 已覆盖统一异常 catch、MovePictures unknown 和 TryAddPicture 参数失败。

### FIX-007：CLOSED

- `decisions.md`、`exception-contract.md`、`date-contract.md` 和 docs 已同步；Docs.Tests 与 PackageConsumer 的职责描述不再混淆。
- 最终报告计数已更新为三个 Unit TFM 各 470/471，Integration net6/net8 各 15/15，Docs 10/10。

## 最终验证证据

- 冻结后 Release build：0 errors / 28 warnings。
- Unit：netcoreapp3.1、net6.0、net8.0 各 470 passed / 1 failed / 0 skipped；唯一失败为 formal API hash。TRX：`artifacts/test-results/unit/rc-final-*`。
- Integration：net6.0、net8.0 各 15/15；Docs：10/10。TRX：`artifacts/test-results/integration/rc-final-*`、`artifacts/test-results/docs/rc-final-*`。
- DateTimeOffset 双 TZ：各 6/6；Failure Workbook Diagnostics 独立复跑 3/3。
- PackageConsumer：netstandard2.0 编译通过；三个 runtime 均运行通过；五个包内 DLL 与冻结输出一致。
- 弃用扫描未发现旧 Mapping builder、runtime v1 migration 或 DataTable CsvHelper 的可编译残留；七个 NPOI 扩展容器仍为 public；生产程序集间没有新增 IVT。
- `git diff --check` 无 whitespace error，仅有 ProfileFixtures XML 的 CRLF/LF warning。

## 发布结论

当前必须保持 `No-Go`：开放 P1 是正式 APICompat/审批和性能资源预算/完整证据，两者均属于明确发布门禁。修复或取得批准后，应重跑对应 API、Benchmark/Resource 与最终全量验证，再执行一次独立复审；只有开放 P0/P1 为零时才能考虑 Go。

