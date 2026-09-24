# Benchmark Report

## Result

`PARTIAL`：Benchmark 项目在 net8 Release 编译通过，`ProviderComparisonProbe` 和 `ProviderComparisonBenchmarks` 已加入 ClosedXML 分支，支持 `closedxml` provider filter。首次运行发现 Benchmark 输出缺少 ClosedXML 传递依赖，已将 `ClosedXML [0.105.1]` 作为 Benchmark 项目的显式运行时包引用；修复后重新构建并完成三档真实 Probe。

```text
dotnet build benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj --no-restore -v:minimal
```

采样命令（每档同步/外围异步各 3 次，包含一次 warmup，隔离 worker process）：

```text
dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build -- --provider-comparison-probe <artifact> <rows> 3 after closedxml
```

| Rows | Mode | Median elapsed | Median allocated | Median peak WS | Output bytes | Median rows/s | Evidence |
| ---: | --- | ---: | ---: | ---: | ---: | ---: | --- |
| 1,000 | sync | 119.23 ms | 12,721,664 B | 106,582,016 B | 28,773 | 8,386.93 | `artifacts/closedxml-1k.jsonl` |
| 1,000 | async | 118.42 ms | 12,861,728 B | 106,655,744 B | 28,773 | 8,444.27 | same |
| 10,000 | sync | 774.18 ms | 121,066,200 B | 135,049,216 B | 223,500 | 12,916.84 | `artifacts/closedxml-10k.jsonl` |
| 10,000 | async | 813.06 ms | 121,945,712 B | 135,536,640 B | 223,500 | 12,299.21 | same |
| 100,000 | sync | 6,033.16 ms | 1,166,304,664 B | 340,566,016 B | 2,221,652 | 16,575.07 | `artifacts/closedxml-100k.jsonl` |
| 100,000 | async | 3,863.07 ms | 1,173,771,040 B | 334,151,680 B | 2,221,652 | 25,886.15 | same |

运行环境为 Windows `win-x64`、.NET 8 Release、当前 dirty worktree；上表 workload 是相同的普通标量 List/Workbook export + import roundtrip。`before=NOT_APPLICABLE`，没有历史 ClosedXML 等价实现，不进行 before/after 伪比较。

## ClosedXML 专项场景

新增 `ClosedXmlScenarioProbe`，以相同实体行和重复次数分别测量独立导出、独立导入、真实文件提交/读回，以及多 Sheet、样式、模板、公式 round-trip。每个场景均有 sync/外围 async 各 3 次（含 1 次 warmup），完整样本、GC 和峰值工作集见对应 `artifacts/closedxml-scenarios-*.jsonl`；下表为中位数，字节列为工作簿输出或导入 payload 大小。

```text
dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build -- --closedxml-scenario-probe <artifact> <rows> 3 all
```

| Rows | Scenario | Operation | Sync ms | Async ms | Sync alloc | Async alloc | Bytes |
| ---: | --- | --- | ---: | ---: | ---: | ---: | ---: |
| 1,000 | export | export | 67.17 | 54.77 | 7,047,736 | 7,079,280 | 31,456 |
| 1,000 | import | import | 124.93 | 124.10 | 8,506,096 | 8,619,080 | 31,456 |
| 1,000 | file | file | 222.22 | 209.15 | 15,562,088 | 15,668,608 | 31,456 |
| 1,000 | multisheet | roundtrip | 329.00 | 342.23 | 26,318,928 | 26,507,328 | 52,378 |
| 1,000 | style | roundtrip | 305.84 | 315.58 | 19,931,152 | 20,076,768 | 31,591 |
| 1,000 | template | roundtrip | 236.02 | 162.86 | 16,844,256 | 16,945,496 | 33,315 |
| 1,000 | formula | roundtrip | 254.90 | 413.93 | 15,796,136 | 15,946,744 | 34,048 |
| 10,000 | export | export | 351.50 | 259.97 | 67,159,264 | 67,383,776 | 251,413 |
| 10,000 | import | import | 912.86 | 340.77 | 77,969,752 | 78,145,896 | 251,413 |
| 10,000 | file | file | 538.02 | 622.47 | 144,635,664 | 145,280,936 | 251,413 |
| 10,000 | multisheet | roundtrip | 1,106.30 | 1,224.89 | 248,950,072 | 250,709,008 | 448,775 |
| 10,000 | style | roundtrip | 1,186.50 | 1,293.48 | 187,989,024 | 188,897,168 | 251,548 |
| 10,000 | template | roundtrip | 672.58 | 780.88 | 146,952,544 | 147,865,632 | 257,168 |
| 10,000 | formula | roundtrip | 878.55 | 924.22 | 147,585,976 | 148,514,216 | 272,945 |
| 100,000 | export | export | 2,661.55 | 2,122.24 | 648,072,576 | 650,548,120 | 2,504,726 |
| 100,000 | import | import | 5,588.51 | 3,891.46 | 766,471,248 | 771,714,856 | 2,504,726 |
| 100,000 | file | file | 8,818.46 | 6,200.28 | 1,414,573,432 | 1,419,757,624 | 2,504,726 |
| 100,000 | multisheet | roundtrip | 12,263.04 | 12,658.61 | 2,443,908,216 | 2,458,907,112 | 4,510,390 |
| 100,000 | style | roundtrip | 9,315.82 | 8,944.45 | 1,847,723,592 | 1,855,488,296 | 2,504,904 |
| 100,000 | template | roundtrip | 13,342.09 | 7,979.04 | 1,426,844,344 | 1,434,599,064 | 2,512,300 |
| 100,000 | formula | roundtrip | 6,794.06 | 10,503.95 | 1,448,981,912 | 1,462,190,016 | 2,721,726 |

100K 富场景的峰值工作集约 422 MiB，file async 样本约 348 MiB；这些数字是当前 Windows/.NET 8 机器的观测值，不是跨平台容量承诺。500K/1M、跨 OS/字体/AutoFit 峰值仍为 `BLOCKED_APPROVAL` 或 `BLOCKED_EXTERNAL`。
