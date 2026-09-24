# Benchmark Plan

`ProviderComparisonBenchmarks` 与 `ProviderComparisonProbe` 已加入 ClosedXML，固定与 NPOI/MiniExcel 相同的行数据、请求、sync/async 模式和 1K/10K/100K 参数。另有 `ClosedXmlScenarioProbe` 独立测量 `export`、`import`、`file`、`multisheet`、`style`、`template`、`formula`，每个场景固定 3 次重复和一次 warmup，并记录 elapsed、allocated bytes、GC、峰值工作集、输入/输出字节和吞吐。ClosedXML 没有历史 before 等价数据，因此 before 结果记录 `NOT_APPLICABLE`，不伪造 before/after 差异。

已完成 Windows/.NET 8 Release 的 1K/10K/100K 普通 round-trip 和专项场景采样，原始 JSONL 及中位数见 `benchmark-report.md`。500K/1M、跨 OS/字体/AutoFit 和生产机器峰值需要批准资源预算或外部环境后再运行。
