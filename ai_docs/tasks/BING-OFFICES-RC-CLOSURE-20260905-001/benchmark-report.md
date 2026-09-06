# Benchmark 报告

## 当前结论

状态：`PARTIAL / budget unapproved`。已完成泛型 Sheet 委托缓存，以及 Excel、CSV、Failure Workbook 1k/10k/100k 真实公共调用链 ShortRun。没有历史可比 baseline 和维护者批准预算；模板/样式/验证/图片专项未在本轮完整运行，不使用“低 GC”结论。

## 环境

- Commit：`9d78ab76e28891ac9b7e0e558f4df669721a4ea6` + 当前 dirty diff
- OS：Windows 10 10.0.19045.6466
- CPU：Intel Core Ultra 7 270K Plus，24 logical/physical cores
- SDK：10.0.400
- Runtime：.NET 8.0.30，x64 RyuJIT AVX2
- GC：Concurrent Workstation
- BenchmarkDotNet：0.14.0
- Job：ShortRun，1 launch，3 warmup，3 iterations
- Diagnoser：MemoryDiagnoser

## 泛型 Sheet 分派

| Method | Sheets | Mean | StdDev | Gen0 | Gen1 | Allocated |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Export | 1 | 1.265 ms | 0.2446 ms | 68.3594 | 33.2031 | 1265.20 KB |
| Import | 1 | 1.106 ms | 0.0020 ms | 35.1563 | 7.8125 | 668.55 KB |
| Export | 8 | 2.569 ms | 0.3614 ms | 93.7500 | 42.9688 | 1787.59 KB |
| Import | 8 | 2.242 ms | 0.0274 ms | 66.4063 | 19.5313 | 1285.32 KB |

场景通过公开 `IExcelExporter`/`IExcelImporter`、DI 和真实 XLSX 序列化/读取执行。无旧反射版本原始 artifact，因此 Ratio/改善百分比不可用。Import 运行记录内部 exception 计数（1 Sheet 3、8 Sheet 10），未逃逸到调用方且进程 exit 0；需在完整性能审查中解释，不据此判定异常。

Export 的 3 次迭代方差较高，99.9% CI 不稳定；本 ShortRun 证明场景可执行，不作为批准预算。

原始 artifacts：`artifacts/benchmark/generic-dispatch-shortrun/`；Dry 有效性探针：`artifacts/benchmark-dry/generic-dispatch-network/`。

## Excel Stream 1k/10k/100k

| Method | Rows | Mean | StdDev | Gen0 | Gen1 | Gen2 | Allocated |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Import | 1,000 | 25.97 ms | 2.682 ms | 888.8889 | 444.4444 | - | 17.80 MB |
| Export | 1,000 | 15.12 ms | 4.876 ms | 500.0000 | 406.2500 | 218.7500 | 8.36 MB |
| Import | 10,000 | 320.83 ms | 33.635 ms | 10,000 | 5,000 | 2,000 | 168.74 MB |
| Export | 10,000 | 119.16 ms | 2.813 ms | 5,000 | 4,000 | 1,666.6667 | 71.12 MB |
| Import | 100,000 | 2.074 s | 128.384 ms | 95,000 | 49,000 | 6,000 | 1,684.74 MB |
| Export | 100,000 | 1.297 s | 4.080 ms | 36,000 | 30,000 | 4,000 | 734.51 MB |

`ExportDestinationCapacity` 同样完成三档，100k 为 1.250 s / 734.51 MB。Import 记录每 benchmark 三次内部 exception 事件但进程均 exit 0；需进一步定位 BenchmarkDotNet 事件含义，不把它解释为业务失败。原始 artifacts：`artifacts/benchmark/stream-pipeline-shortrun/`。

## CSV 1k/10k/100k

| Method | Rows | Mean | StdDev | Gen0 | Gen1 | Gen2 | Allocated |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Import | 1,000 | 736.0 us | 48.61 us | 109.3750 | 35.1563 | - | 2.03 MB |
| Export | 1,000 | 1.834 ms | 471.04 us | 812.5000 | 210.9375 | - | 14.62 MB |
| Import | 10,000 | 9.396 ms | 209.53 us | 1,171.8750 | 406.2500 | 156.2500 | 20.09 MB |
| Export | 10,000 | 12.178 ms | 43.13 us | 8,000 | 703.1250 | 281.2500 | 144.42 MB |
| Import | 100,000 | 89.726 ms | 838.16 us | 11,000 | 3,500 | 500 | 198.32 MB |
| Export | 100,000 | 124.293 ms | 2.567 ms | 80,200 | 4,400 | 400 | 1,448.26 MB |

场景通过 DI 解析公开 `ICsvImporter/ICsvExporter`，使用 typed 行和 DateTimeOffset `O` 文本。原始 artifacts：`artifacts/benchmark/csv-pipeline-shortrun/`。

## Failure Workbook 1k/10k/100k

| Failure rows | Mean | StdDev | Gen0 | Gen1 | Gen2 | Allocated |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 1,000 | 72.34 ms | 0.993 ms | 1,875 | 1,375 | 625 | 28.54 MB |
| 10,000 | 367.64 ms | 42.761 ms | 14,000 | 9,000 | 4,000 | 263.90 MB |
| 100,000 | 4.115 s | 121.691 ms | 132,000 | 71,000 | 9,000 | 2,606.14 MB |

场景使用真实 `ErrorRowsOnly` 导入与失败工作簿序列化。100k 的高分配和 Gen2 是明确风险，不符合“低 GC”描述。原始 artifacts：`artifacts/benchmark/failure-workbook-shortrun/`。

## 命令

```powershell
dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build --no-restore -- --filter '*GenericSheetDispatch*' --job short --artifacts ai_docs/tasks/BING-OFFICES-RC-CLOSURE-20260905-001/artifacts/benchmark/generic-dispatch-shortrun --exporters json markdown csv
dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build --no-restore -- --filter '*StreamPipelineBenchmarks*' --job short --artifacts <task>/artifacts/benchmark/stream-pipeline-shortrun --exporters json markdown csv
dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build --no-restore -- --filter '*CsvPipelineBenchmarks*' --job short --artifacts <task>/artifacts/benchmark/csv-pipeline-shortrun --exporters json markdown csv
dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build --no-restore -- --filter '*FailureWorkbookBenchmarks*' --job short --artifacts <task>/artifacts/benchmark/failure-workbook-shortrun --exporters json markdown csv
```
