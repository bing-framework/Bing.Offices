# 性能基线

## 当前状态

已完成 BenchmarkDotNet Dry smoke，并建立了固定迭代基线配置。MappingValidation 基准现已覆盖真实 named validator、计划 cold/cache hit/cache miss、JSON/XML v1/v2、Unique journal、Profile 注册和 cache eviction；Dry smoke 执行 208 个基准且无运行时异常。

当前观察：cache hit 的 `DynamicPlanBuildCacheHit` 分配显著低于 cold/miss 场景；Dry job 报告多数场景单次迭代时间低于推荐 100ms，因此该结果只能作为 smoke/结构证据，不能作为发布性能阈值。完整基线仍需固定 Job、输入规模和输出采集格式后执行。

产品路径与资源压力已分离：MappingValidation 不再主动分配 90KB byte[]；LOH/working set 压力保留在独立 ResourceProbe，输出使用 `lohSampledPeakBytes` 和 `lohRetainedBytes`，不将假负载称为产品管线 LOH。

固定基线配置为 `launchCount=1`、`warmupCount=2`、`iterationCount=3`，由两个基准类共享，且不绑定本机 SDK moniker；当前 net8.0 目标实际运行在 `.NET 8.0.27`，主机 SDK 为 `10.0.300`。代表性 `DynamicPlanBuildCacheHit` 已执行全部 16 个参数组合：Mean 范围约 `262.8 us` 至 `1.4048 ms`，Allocated 范围约 `585.24 KB` 至 `2,949.93 KB`，同时采集 Gen0/Gen1。原始报告位于 `BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.MappingValidationBenchmarks-report-github.md`。

该基线是单机、3 次测量的比较起点，不是发布性能阈值；后续优化必须在相同 Job、输入规模和运行时环境下重跑。Stream import/export smoke 已执行，但完整、可归因的前后优化对比仍未完成。

Round 4 已将 Plan cache key 的序列化改为直接 UTF-8 字节路径，保留 SHA256/Base64 输出契约；Round 5 已将 Importer/Exporter 的 workbook plan 分组和泛型反射调度移出主类；Round 6 新增同 payload 的旧/新序列化对照基准。固定 Job（`launchCount=1`、`warmupCount=2`、`iterationCount=3`）下，代表性组合中旧式 `Serialize` + `Encoding.UTF8.GetBytes` 约 `5.89 KB`、`2.3 us`，直接 `SerializeToUtf8Bytes` 约 `2.52 KB`、`2.15 us`；64 个参数组合均执行完成。由于 Dry 基准单次迭代低于推荐 100ms，该结果作为可归因结构/分配证据，不作为发布性能阈值。

未完成的产品性能风险仍包括：导入 source 到 `MemoryStream` 的整块复制、失败工作簿 `ToArray()`、NPOI DOM 的工作簿内存占用、完整 Stream import/export 前后对比，以及更深层行处理、单元格写入和 Loader 解析职责拆分。

## 约束

- 在正确性和 Mapping 合并边界完成前不进行性能重写。
- 后续 benchmark 必须记录输入规模、输出大小、Mean/Median（可用时）、Allocated、Gen0/1/2、LOH、managed peak 和 working set。
- 不把主动资源压力探针描述为产品路径的 LOH 证据。
