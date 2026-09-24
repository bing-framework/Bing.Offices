# Bing.Offices ClosedXML Provider 实施计划

## 任务信息

- Task ID: `BING-OFFICES-CLOSEDXML-PROVIDER-20260922-001`
- 执行模式: execute-plan
- Provider: `Bing.Offices.ClosedXml`
- 目标 TFM: `net6.0;net8.0`
- ClosedXML: `0.105.1`（执行时如稳定版本变化，必须在 decisions.md 记录）

## 目标与边界

新增与 NPOI、MiniExcel 同等级的 ClosedXML Provider，复用现有 Workbook Request、Mapping Plan、Converter、Validation、Relations、异常、资源限制和原子文件提交契约。ClosedXML 类型只能出现在 Provider 程序集内部。

第一版支持：XLSX List/Workbook、多 Sheet、Mapping、Dynamic Column、Converter、Validation、Relations、基础 Style、Merge、Formula 保存/读写、Template、Entity Layout、真实外围 Async/Cancellation、Stream ownership、Atomic File Commit 和资源预检。

第一版不支持：XLS、Chart 创建、PivotTable、宏执行、未经 Spike 证明的 XLSM 宏保留、完整 Excel Formula Engine、静默丢弃不支持结构、ClosedXML DOM 并发访问、Task.Run/.Result/.Wait 伪异步。

Template 只承诺已验证的 Worksheet、Cell Value、Style、Merge、Formula、Comment、Validation、Conditional Formatting；无法可靠保留的 Chart/Image/Macro/PivotTable 必须在 `XLWorkbook` 创建前 fail-fast。

## 实施顺序

1. `CX-000..003`: 冻结 dirty worktree 候选，验证 ClosedXML API、依赖、TFM、Stream、Formula、Style、Template、XLSM、Chart、Image、Table、AutoFilter、字体和非线程安全边界，写入 baseline/decisions/matrices。
2. `CX-100..104`: 盘点公共 requirement，保持既有 capability bit 数值，只有有公共语义的能力才追加细粒度 bit；更新 API Snapshot、第三方 Provider Consumer 和 fail-fast 规则。
3. `CX-200..203`: 新建 Provider csproj/AssemblyInfo，接入 solution、`ClosedXmlPackageVersion`、依赖图和双 TFM 构建；不修改 Bing.Offices 版本号。
4. `CX-300..306`: 实现 ClosedXML Adapter、Mapping/Value/Style/Workbook/Preflight/Async staging 内部职责和 `AddBingOfficesClosedXml`，保持 first-registration-wins。
5. `CX-400..405`: 实现真实 XLSX Export；所有请求经 Core Plan -> capability preflight -> ClosedXML -> SaveAs -> reopen/read-back。
6. `CX-500..506`: 实现真实 XLSX Import；覆盖 CellKind、日期序列、公式/缓存值、空白、错误、合并、隐藏行列、重复列、校验、动态列和 Relations。
7. `CX-600..606`: 实现 Style、Merge、Template、Entity Layout；Chart 创建保持 Unsupported；不支持的模板结构在 DOM 创建前拒绝。
8. `CX-700..706`: 实现 ZIP/XML/资源预检、DOM admission、取消、临时文件清理、原子提交和真实外围异步；ClosedXML DOM 串行化。
9. `CX-800..804`: 新增 ClosedXML Unit/Integration，核心行为使用真实 XLWorkbook/MemoryStream/FileStream/XLSX，补齐映射缓存、资源、取消、流所有权和异常测试。
10. `CX-900..904`: 接入 Cross-provider、ThirdParty Provider Consumer、net6/net8 Package Consumer、API Snapshot 和公共文档。
11. `CX-1000..1004`: 接入真实 IO Benchmark/Resource Probe，固定数据集、TFM、Runtime、Seed、指标；新能力无历史 before 时标记 `NOT_APPLICABLE`，超时/拒绝记录边界，不伪造统计。
12. `CX-1100..1103`: 更新 README、Excel Provider 矩阵、进度、示例和最终门禁报告。

## 必须生成的任务文档

Create/执行阶段维护：

```text
baseline.md
requirements-matrix.md
provider-capability-matrix.md
api-diff.md
test-matrix.md
benchmark-plan.md
decisions.md
execution.md
```

执行结果报告另行生成，不以计划代替真实证据。

## 验收与门禁

- Core/Abstractions、NPOI、MiniExcel 旧测试无回归。
- ClosedXML 双 TFM Build、Unit、Integration、Cross-provider、Package Consumer 通过。
- Unsupported Feature 在 `XLWorkbook` 创建前失败，异常包含 Provider/Operation/Stage/Feature。
- Resource limits 在 DOM 创建前生效。
- Cancellation、Stream ownership、Dispose、Atomic File Commit 有真实文件证据。
- ClosedXML 类型不泄漏到公共 API；已有 capability bit 不改变；未批准 Breaking Change 为 0。
- Benchmark、Resource、外部 CI、人工审批分别记录真实状态，不将 BLOCKED/NOT_APPLICABLE 当作代码 PASS。

## 偏差规则

若实际 ClosedXML API 与本计划不同，以稳定包和真实 Spike 为准，记录 `decisions.md`，同时更新 capability matrix、requirements matrix、test matrix 和 execution.md。不得静默降级或伪造能力。
