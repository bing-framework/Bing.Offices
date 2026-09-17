# 最终报告

- Task ID：`BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001`
- 分支：`feat/miniexcel-provider`
- 当前状态：`PARTIAL`
- 独立审查：当前 `review.md` Round 2 仍为 `NEEDS_FIX`；本轮已完成 FIX-004、FIX-007、FIX-008、FIX-009，等待下一轮独立复审
- 发布结论：当前不可标记为已发布；需补齐外部/审批证据。

## 1. Preflight 与 netstandard2.0

实际复核没有复现历史报告所称的 `ExcelXlsxZipPreflight` netstandard2.0 编译错误；Core Release 构建为 0 warning/0 error。最终方案保留 Core 的 ZIP/XML 安全策略，将同一源文件链接编译到 NPOI 和 MiniExcel Provider，从而不升 Core TFM、不复制两套算法、不依赖生产 IVT。路径、重复 entry、entry/总量/压缩比、DTD、XML 深度/字符预算、取消和流位置语义保持不变。

## 2. 生产 IVT

Core 到 `Bing.Offices.Npoi` 和 `Bing.Offices.MiniExcel` 的生产 `InternalsVisibleTo` 已为 0。各 Provider 仅保留测试友元；没有引入公共 SPI。`PublicApiContractTest` 和 API identity/compare 均通过。

## 3. 日期与 RowIndex

MiniExcel 日期适配现在复用 Core `DateTimeExcelValidationRule`/`ExcelDateParser`，传递 workbook `date1904`；第三方已返回 `DateTime` 时在 OLE Automation 范围内恢复 serial，并由 Core 统一处理 1900/1904 边界，覆盖真实 1900/1904 serial、DateTime、DateTimeOffset 和跨目标原始值；无 offset 的 raw/serial 值遵守 `RequireExplicitOffset`。固定列、动态列 converter 都收到从数据行 2 开始的 1-based `RowIndex`，固定/动态 `ColumnIndex` 均使用物理表头位置。MiniExcel 专项双 TFM 25/25、跨 Provider 合同双 TFM 5/5。

## 4. 性能与关系

100K probe（schema 2，`.NET 8.0.30`）结果：Sync `1972.34 ms / 1,276,698,832 B / 102/13/3 GC / 111,501,312 B working set / 50,701 rows/s / 2,078,325 B output`；Async `1258.70 ms / 1,275,428,960 B / 100/11/1 GC / 133,664,768 B working set / 79,447 rows/s / 2,078,320 B output`。另有同 workload NPOI/MiniExcel 对照：`artifacts/review-fix/provider-comparison-100k-final.json`。elapsed/GC/working set 受机器状态影响，当前数据不作为跨运行稳定基线；没有可重放的修改前样本，因此优化收益为 `NOT_VERIFIED`。关系绑定保持最坏 `O(P×C)`，没有在缺少稳定收益证据时重写 comparer 或重复键语义。

500K、1M、生产 2 CPU/4 GiB 和外部 CI：`NOT_VERIFIED`。

## 5. 产物与残留

正式证据统一位于根 `artifacts/`：

- `artifacts/tests/`：TRX 测试结果
- `artifacts/packages/`：四个 nupkg 和四个 snupkg
- `artifacts/consumers/review-round2-final/`：最终包 PackageReference Consumer 独立 cache、build/run 日志和输出
- `artifacts/benchmarks/`：Benchmark 列表和 MiniExcel 100K probe
- `artifacts/review-fix/`：本轮 NPOI/MiniExcel 同 workload probe
- `artifacts/api-round2-final3/`：最终包 API capture/compare、identity self-test 和篡改拒绝证据
- `artifacts/resource-probe/`：资源矩阵 JSONL/Markdown
- `artifacts/api/`：既有 capture、compare 和 identity 日志

`src`、`tests`、`benchmarks` 下 `artifacts`、`TestResults`、`BenchmarkDotNet.Artifacts`、`packages` 残留扫描计数为 0。前一任务包已迁移到 `artifacts/legacy-project-packages/`。

## 6. 中文注释

按 `chinese-comments` 规则复核了本任务实际修改的 Core preflight、日期 parser、NPOI/MiniExcel Provider、Importer/Exporter/adapter、测试和 Benchmark 代码；新增复杂行为有中文 XML/方法体说明，公共实现优先使用 `<inheritdoc />`。未对全仓库既有内部辅助成员做无关批量治理，后续缺口不影响本任务签名或行为。

## 7. API、Consumer 与发布阻塞

公共 API 无 diff，baseline 只同步最终候选源码和包 identity 元数据。最终 `artifacts/packages/` 四个包的条目与 `artifacts/consumers/review-round2-final/` net6/net8 输出 DLL SHA-256 逐项相符；PackageReference restore/build/run 均通过，MiniExcel nuspec 无 NPOI 依赖。最终包内 DLL 替换为文本后 API compare 以退出码 `1` 拒绝。资源矩阵 36/36 通过但 `approvalStatus=BLOCKED`，需要批准人/时间；500K/1M、生产机器和外部 CI 仍为 `NOT_VERIFIED`，这些是当前发布阻塞项。

## 8. 版本与 Git

`version.props` 和 `version.dev.props` 结束时 SHA-256 与任务起点一致：

- `version.props`：`77FEBEC429F841DC839F84FED57A48AE4159DF2EA92A15AD808BB3027C39B0D1`
- `version.dev.props`：`BABFE7703A15E1C11F46B45D34BC59D7913ECB1E3DCFDA901E72F4D60A8D134D`

DLL/NuGet identity 均为既有 `2.0.0`，没有修改版本文件或版本属性。未执行 commit、push、tag、PR 或 NuGet publish。

## 参考报告

- [执行报告](execution.md)
- [独立审查](review.md)
- [API Diff](api-diff.md)
- [Unit 测试报告](unit-test-report.md)
- [Integration 测试报告](integration-test-report.md)
- [Package Consumer 报告](package-consumer-report.md)
- [Benchmark 报告](benchmark-report.md)
- [Resource 报告](resource-report.md)
- [Symbol-Test 映射](symbol-test-map.md)
