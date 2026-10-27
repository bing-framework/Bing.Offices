# 低内存容量验收

当前证据由 Review Round 1 重测的 4 个导出场景和复用的 4 个导入场景组成。每个场景独立进程、10 列，数据量仅 10,000／100,000 行，不设置固定 MB 门槛。导出临时峰值已覆盖暂存包与提交副本，原始旧结果保留，不覆盖历史证据。

本轮采集环境：Windows 10 专业工作站版 19045，OS/进程 X64，.NET 8.0.30，NPOI 2.7.4；Intel Core Ultra 7 270K Plus，24 核/24 逻辑处理器，物理内存 34,080,526,336 字节（约 31.74 GiB）。工作目录 H: 为 ZHITAI TiPlus7100s 2TB / NVMe / NTFS，预检可用约 1.41 TiB；系统临时目录位于 C:，CT1000E100SSD8 / NVMe / NTFS，预检可用约 122 GiB。导入旧 JSON 没有 CPU/磁盘字段，本轮未修改旧 JSON 补写历史机器信息。

|方向|数据|行数|耗时 s|行/s|首行 ms|进程峰值 MiB|库临时峰值 MiB|NPOI 临时峰值 MiB|落盘|残留|
|---|---|---:|---:|---:|---:|---:|---:|---:|---|---:|
|export|repeated|10000|1.323|7561|687.4|92.4|0.60|0.30|False|0|
|export|repeated|100000|3.084|32421|305.1|96.3|5.89|2.95|False|0|
|export|unique|10000|0.954|10480|312.9|92.3|0.71|0.36|False|0|
|export|unique|100000|3.256|30714|311.3|96.2|7.06|3.53|False|0|
|import|repeated|10000|0.442|22612|190.7|72.1|0.28|0.00|False|0|
|import|repeated|100000|1.820|54935|194.6|80.7|2.76|0.00|False|0|
|import|unique|10000|1.048|9543|361.2|98.9|3.42|0.00|True|0|
|import|unique|100000|8.168|12242|3428.3|112.6|34.16|0.00|True|0|

8 个场景均有 success 证据，导出全量读回校验、导入逐行逐列校验通过。两个 unique 导入场景实际观察到输入、数据和 12 字节定长记录索引文件，确认超过默认 8 MiB 预算并落盘。导入负载不含日期、错误单元格或停止阈值交叉，本轮对应行为修复不影响这些负载的暂存、索引及文本读取路径，因此复用原始结果。

4 个新导出 JSON 均有 `commitOverlapObserved=true`。repeated 10k/100k 的提交副本峰值分别为 314,022 / 3,088,130 字节，unique 10k/100k 分别为 372,628 / 3,701,520 字节；库总峰值为对应两份包的总和。session、提交文件及 NPOI 临时文件分别检查，残留均为 0。旧导出表遗漏提交文件，只能作为历史结果，不能用于当前总额度证明。

测量从 API 调用开始到返回，覆盖输入暂存、共享字符串索引、逐行处理、最终打包、文件提交和清理。导入 fixture 生成与导出读回校验在测量外。进程峰值在 API 返回后读取 Process.PeakWorkingSet64，包含进程初始化；工作集、托管内存、NPOI 临时盘及导入临时盘以 20 ms 采样，可能漏过极短峰值。新导出还在每次真实提交 Write/WriteAsync 完成后采样库临时文件，以活动文件流 Length 补足元数据尚未反映的缓冲字节；不强制 Flush。采样装饰器仍委派默认原子提交器，其观测开销包含在导出耗时内，不据此推导与旧探针的精确性能增减。报告中的导出 firstRowMs 指惰性源首行被枚举，导入指消费者收到首行；不是可下载文件首字节时间。

NPOI 自建临时文件按 poifiles 目录本次新增文件观测，不计入库 512 MiB 硬额度。并行进程可能干扰该目录观测，本轮场景串行执行。结果依赖本机磁盘、运行时与数据，不代表内存 SLA。

## 执行保护与未验证指标

4 格预期均为 EXPECTED_SUCCESS。首次 repeated/100k 在 API 已产出文件及清理之后，由后续 DOM 全量校验触发本地 1 GiB 进程保护并受控终止，未记为通过；保留 `run-summary-first-attempt.json`。明确将包含 DOM 校验的进程保护改为 4 GiB 后补跑该格及剩余 unique 两格，90 秒超时和 1 GiB 临时盘保护不变。repeated/10k 直接复用已通过结果；最终 4 格均成功。保护预算不作为产品容量门槛。

|指标|状态|证据边界|
|---|---|---|
|完整 API 耗时、首行延迟、进程峰值、采样托管堆、临时文件峰值/清理|VERIFIED|本表及 JSON；托管堆是存活量采样，不等于总分配量|
|DOM 与新模式同负载性能对照|NOT_VERIFIED|导出后的 DOM 读回只验证正确性，不是性能对照|
|累计分配量、GC 次数/暂停时间|NOT_VERIFIED|当前 JSON 没有相应指标，不能由托管堆峰值推导|
|取消请求到释放完成的延迟|NOT_VERIFIED|已有取消与目标保护正确性测试，不构成延迟测量|
|固定 MB/SLA、50 万/100 万行|NOT_APPLICABLE|不属于本次批准验收范围|

因此批准的 10k/100k 正确性和资源观测已通过；原方案扩展性能指标仍未验证，不能笼统称所有性能项 PASS。

复现命令（替换方向、数据类型和行数，所有组合运行一次）：

```powershell
dotnet build tests/Bing.Offices.ResourceProbe/Bing.Offices.ResourceProbe.csproj -c Release
dotnet tests/Bing.Offices.ResourceProbe/bin/Release/net8.0/Bing.Offices.ResourceProbe.dll --streaming-scenario artifacts/low-memory/resources export repeated 10000
```

历史 8 份 JSON 保存在 `resources/`；当前 4 份导出 JSON 和执行保护记录在 `resources/fix-round1/`。实际受控运行脚本及日志在 `artifacts/low-memory/fix-round1/`。代码或数据分布未变化时，不因报告编辑重复测量。
