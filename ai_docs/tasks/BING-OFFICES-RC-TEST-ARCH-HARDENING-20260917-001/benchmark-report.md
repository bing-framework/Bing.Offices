# 性能与热点报告

- MiniExcel 100K：`artifacts/benchmarks/rc-test-arch-mini-excel-100k.json`，sync 3551.49 ms，async 2820.00 ms；保留为现有 Provider 探针，不重标为正式发布性能结论。
- 历史 Provider comparison：`artifacts/benchmarks/rc-test-arch-provider-after-100k.json` 保持原样，仅作为旧 after/WorkingSet 样本。
- 本轮 Provider comparison：`artifacts/benchmarks/review-fix-provider-npoi-100k.json` 与 `review-fix-provider-miniexcel-100k.json`。每个 Provider 的 sync/async 模式由独立子进程运行，warmup 后各 3 次；首行含 `processId`、`providerFilter`、`providerCount=1` 和候选程序集 SHA-256，峰值使用 `Process.PeakWorkingSet64`。由于没有同一候选的 pre-change Provider 对照，不能计算 Provider 间 before/after 收益。
- Real IO：`artifacts/benchmarks/rc-test-arch-real-io.json`，1K/10K/100K CSV/Excel sync/async 及受控异步共 24 scenarios × 3 repetitions，`status=passed`。

## P9-P11 Hotspot

Round 4 的 RawDate/Relation JSONL 仍作为调查资料保留，但不再是当前候选的性能验收证据。RawDate 的“before”是简化 numeric-cell replay，不是 HEAD Reader；Relation 的“after”来自已撤回的父键缓存/索引实现。Round 7 已恢复 `MiniExcelRawDateSerialReader` 的 HEAD 全量索引和单次 Query 调用，撤回未经可信 baseline 支持的 RawDate 过滤/重复 Query 优化。当前热点入口仅用于可运行的全量 numeric-cell 调查 workload，不代表生产性能收益。

因此本轮不报告 RawDate 或 Relation 的 before/after 收益，也不把旧 Relation after 的 0.9360 ms 写成当前结果。原始文件和候选 identity 保持不变，禁止重标。LOH 字段仅表示测前强制 GC 后与测后最近一次 GC 的 `GenerationInfo[3].SizeAfterBytes` 快照较大值，不是两侧完整 GC 后值，也不是 workload 峰值；PeakWorkingSet 才使用 `Process.PeakWorkingSet64`。

Importer 的固定/动态物理列预绑定已随未证明优化一并撤回；`ReadColumnRange` 的绝对列号语义由直接回归覆盖，未声明性能收益。Dictionary 复制、setter 和关系扫描不作本轮优化。后续若要重新优化，必须先用真实原实现建立同 workload、同候选 identity 的独立三次 before/after，并按字段分别取中位数。

## 结论

P9/P10/P11 的语义回归和 preflight 边界已保留；RawDate 未证明的过滤优化已撤回，Relation 缓存/索引已撤回。Provider comparison、维护者资源批准、500K/1M、真实 2 CPU/4 GiB、外部候选 CI 和正式 Release 仍为 `NOT_VERIFIED` 或 `BLOCKED`，正式发布状态保持 `PARTIAL`。
