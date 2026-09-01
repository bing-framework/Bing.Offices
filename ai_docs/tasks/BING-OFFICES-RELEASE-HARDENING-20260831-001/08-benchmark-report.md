# Benchmark Report

状态：`PARTIAL`

已完成当前工作树的 BenchmarkDotNet 正式默认作业、独立资源探针和 Round 4 并发尾延迟测量入口。当前结果支持单机代表性 workload 的性能/分配观察；并发结果是带环境身份的可复跑测量，但仍不支持批准的跨机器回归或性能预算结论，也不支持 zero-GC/真实 streaming/完整压缩炸弹防护结论。

## 当前证据

- BenchmarkDotNet 正式过滤集：`StreamPipelineBenchmarks`、`FailureWorkbookBenchmarks`、`MappingValidationBenchmarks`、`DynamicPlanBenchmarks`、`TenantPlanCacheBenchmarks`、`UniqueJournalBenchmarks`，均使用 `DefaultJob` 形成结果摘要。
- 环境：Windows 10.0.19045.6466/22H2，Intel Core Ultra 7 270K Plus，1 CPU，24 logical/24 physical cores，.NET SDK 10.0.300，runtime .NET 8.0.27，X64 RyuJIT AVX2，Concurrent Workstation GC。
- 作业：BenchmarkDotNet `DefaultJob`，由 BDN 自适应 pilot/warmup/actual 采样；不再使用 `LaunchCount=1/WarmupCount=2/IterationCount=3` 的 smoke 属性。
- Stream Pipeline 结果：Import 1K/10K/100K 分别为 `14.736 ms`/`196.018 ms`/`1,956.051 ms`，分配分别为 `16.94 MB`/`162.71 MB`/`1,623 MB`；Export 1K/10K/100K 分别为 `9.536 ms`/`115.328 ms`/`1,338.482 ms`，分配分别为 `8.37 MB`/`71.13 MB`/`734.53 MB`。
- Failure Workbook 结果：1K/10K/100K 分别为 `31.48 ms`/`412.57 ms`/`4,712.89 ms`，分配分别为 `27.55 MB`/`258.63 MB`/`2,554.11 MB`。
- Mapping/缓存结果：动态计划冷构建 100/500 分别为 `23,338.9 us`/`119,067.1 us`；动态计划缓存命中 100/500 分别为 `248.4 us`/`1,252.5 us`；租户计划缓存 100/1000 分别为 `346.6 us`/`11,126.4 us`。MappingValidation 还覆盖 JSON/XML、10K 规则、CacheKey、注册和 working-set 微基准。
- 原始产物目录：任务证据目录 `benchmarks/`；已归档六组基准各自的 `.csv`、`-full-compressed.json`、`-default.md`、`-github.md` 和 `.html`。原始生成目录 `BenchmarkDotNet.Artifacts/results/` 仍保留为本地 BDN 输出，但该目录被 `.gitignore` 忽略，交付引用以任务证据目录为准。
- `UniqueJournalBenchmarks` 结果：唯一列 1/行数 10K、100K 分别为 `3.122 ms`/`52.252 ms`，分配 `6.65 MB`/`64.94 MB`；唯一列 5/行数 10K、100K 分别为 `26.247 ms`/`313.213 ms`，分配 `31.13 MB`/`303.32 MB`；四个场景均包含 Gen0/Gen1/Gen2 和置信区间原始数据。
- 独立资源探针：`artifacts/benchmarks/release-hardening-resource-probe.jsonl`，16 个场景均为 `status=passed`；记录 LOH sampled/retained 与 peak working set ceiling。
- 证据身份：运行时工作树基线为 `1968b24a3ab07b44c3b386a3f761fcdff2fc4315`，同时包含本轮未提交 diff；报告不得被解释为干净 commit 的性能签名。
- 并发尾延迟入口：`dotnet run -c Release --project benchmarks/Bing.Offices.Benchmarks -- --tail-latency artifacts/benchmarks/release-hardening-tail-latency-round4-a.jsonl 256`，固定执行 1/4/16/64 并发；每档先执行最多 64 次预热，再执行 5 轮、每轮 256 次冷 mapping-plan build。worker 先通过 ready gate，再由 start gate 同步开始；延迟从队列提交前时间戳计至 plan 完成，吞吐为每轮墙钟完成量，worker 启动耗时单独记录。每档正式样本为 `1280`，输出 p50/p95/p99、吞吐、重复轮次、样本数和环境身份，`budgetStatus=UNAPPROVED`。
- Round 4 重复证据：`benchmarks/release-hardening-tail-latency-round4-a.jsonl` 与 `benchmarks/release-hardening-tail-latency-round4-b.jsonl`。两次均为 .NET `8.0.27`、Windows `10.0.19045`、24 logical processors、X64、workstation GC，绑定 HEAD `1968b24a3ab07b44c3b86a3f761fcdff2fc4315` 和同一 diff identity。汇总范围如下：并发 1 的 p99 为 `3244-4146 us`、吞吐 `72,809-78,375 ops/s`；并发 4 的 p99 为 `3372-4708 us`、吞吐 `121,961-142,221 ops/s`；并发 16 的 p99 为 `1596-2462 us`、吞吐 `141,041-162,618 ops/s`；并发 64 的 p99 为 `9431-10136 us`、吞吐 `36,141-40,983 ops/s`。高并发波动被保留为事实，不被压缩成 PASS。

## 预算与限制

- 已验证：正式多规模运行、Mean/Error/StdDev、分配量、Gen0/1/2、映射/缓存 workload 和 Failure Workbook workload 均有原始结果。
- 仍不可验证：没有批准的历史 baseline、环境限定预算或具名 release waiver，不能把本轮数值判定为回归通过；并发 1/4/16/64 tail latency 已完成可重复测量，但结果仍标记 `UNAPPROVED`，不能替代性能批准。两次运行在高并发下存在明显 p99/吞吐波动，后续需要版本负责人给出目标 workload、容差和批准人。独立资源探针不等价于 BDN 的 per-operation LOH/working-set 预算。上述限制仍使本报告保持 `PARTIAL`。
- 100K 导入分配约 `1,623 MB`、100K Failure Workbook 分配约 `2,554.11 MB`，与 NPOI DOM 和 staging 行为一致；这明确说明当前实现不是 zero-GC 或真正 streaming。
- `MaxInputBytes` 仅限制复制到 NPOI 前的输入字节数，不限制 ZIP/OLE 解压或 DOM 峰值；资源探针的通过只表示本地设定 ceiling 下场景未越界。
