# 透视缓存客户端与资源探针

这是维护验证工具，不进入生产包，也不要求常规 CI 安装 Office。它通过本地 NuGet 包调用 ClosedXML 或 NPOI Provider，没有生产项目引用。默认 Provider 是 ClosedXML；资源测试与客户端测试是两个独立步骤。

## 构建与生成

F3 验证使用任务产物中的本地 feed、NuGet 配置和隔离缓存。配置位于 `ai_docs/tasks/BING-OFFICES-NPOI-PIVOT-CACHE-20261005-001/artifacts/pivot.NuGet.Config`，包 feed 为同目录下的 `pivot-feed-01`，工具专用缓存为 `pivot-package-cache-01/client-probe`。以下命令显式指定包版本和缓存；不要将工具恢复到未验证的同版本旧包。

```powershell
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$probeTask = Join-Path (Get-Location).Path 'ai_docs/tasks/BING-OFFICES-NPOI-PIVOT-CACHE-20261005-001'
$probeConfig = Join-Path $probeTask 'artifacts/pivot.NuGet.Config'
$probeCache = Join-Path $probeTask 'artifacts/pivot-package-cache-01/client-probe'
$probeOutput = Join-Path (Get-Location).Path ('artifacts/pivot-probe-' + (Get-Date -Format 'yyyyMMddHHmmss'))
New-Item -ItemType Directory -Path $probeOutput | Out-Null

dotnet restore build/PivotClientProbe/PivotClientProbe.csproj --configfile $probeConfig --packages $probeCache -p:BingOfficesPackageVersion=2.0.0
if ($LASTEXITCODE -ne 0) { throw '探针恢复失败。' }
dotnet build build/PivotClientProbe/PivotClientProbe.csproj -c Release --no-restore -m:1 -p:UseSharedCompilation=false -p:BingOfficesPackageVersion=2.0.0
if ($LASTEXITCODE -ne 0) { throw '探针构建失败。' }

dotnet build/PivotClientProbe/bin/Release/net8.0/PivotClientProbe.dll --output (Join-Path $probeOutput 'fixture-net8-async.xlsx') --async true
if ($LASTEXITCODE -ne 0) { throw '夹具生成或完整缓存校验失败。' }
```

配置、feed 和包目录由任务验证生成，不随源码提供；换机器时需重新生成本地包及配置。工具支持 net6.0/net8.0，运行对应 DLL 需要对应运行时。`--provider` 接受 `ClosedXML` 或 `NPOI`；省略时仍使用 ClosedXML。基础透视支持两个 Provider 的同步和异步导出，高级透视也通过相同选择器调用指定 Provider。下面运行 NPOI 基础/高级、net6/net8、同步/异步的完整组合，每种组合使用唯一输出路径：

```powershell
$probeDll6 = Join-Path (Get-Location).Path 'build/PivotClientProbe/bin/Release/net6.0/PivotClientProbe.dll'
$probeDll8 = Join-Path (Get-Location).Path 'build/PivotClientProbe/bin/Release/net8.0/PivotClientProbe.dll'
foreach ($tfm in @('net6.0', 'net8.0')) {
    $probeDll = if ($tfm -eq 'net6.0') { $probeDll6 } else { $probeDll8 }
    foreach ($mode in @('basic', 'advanced')) {
        foreach ($asynchronous in @($false, $true)) {
            $asyncArgument = if ($asynchronous) { 'true' } else { 'false' }
            $probeArgs = @('--output', (Join-Path $probeOutput "npoi-$tfm-$mode-$asynchronous.xlsx"), '--provider', 'NPOI', '--async', $asyncArgument)
            if ($mode -eq 'advanced') { $probeArgs += @('--advanced', 'true') }
            dotnet $probeDll @probeArgs
            if ($LASTEXITCODE -ne 0) { throw "NPOI $tfm $mode async=$asyncArgument 探针失败。" }
        }
    }
}
```

ClosedXML 保持默认选择；将上面的 DLL 分别换成 net6.0 和 net8.0，并用 `--async false`/`--async true`，即可运行基础和高级同步/异步矩阵。也可显式传 `--provider ClosedXML`。`--number-format-probe true` 仅支持 ClosedXML，和上述 NPOI 矩阵分开运行。

这些入口会真实调用 Provider，并验证生成工作簿的完整来源、缓存记录和透视结构。结构校验通过只说明文件结构符合探针断言，不代表 Excel 客户端计算值通过；客户端数值必须由后续 Excel 生命周期步骤单独验证。

默认夹具有 `Data` 和 `Summary` 两个 Sheet，来源为 North 10、South 20、North 30，透视定义使用 Warehouse 行字段及 Amount 求和字段。工具核对完整 Sheet 顺序、表头、全部来源行、每个缓存的字段和全部记录，不用预先写好的 OOXML 代替真实 Provider 输出。成功后产生同名 `.probe.json`。

## 基础 Excel 生命周期

基础验证使用 `verify-excel.ps1`，runner 会通过 `--basic` 让父进程启动并持有专用 Excel 进程，再把 PID 和启动时间传给脚本。脚本验证完整 Data 区域、行字段 Warehouse、Amount 求和字段和完整 `TableRange2`：首次显示 North 40、South 20、总计 60；改首笔金额后记录刷新前显示，显式刷新后验证 45、20、65，并保存、重开只读副本复核来源和结果。脚本直接调用仍兼容，但只有确认新建 Excel 进程不属于启动前已存在的实例后才修改设置或调用 `Quit`；所有权不能确认时不退出该实例。

```powershell
python ./build/PivotClientProbe/run-advanced-excel.py `
    (Join-Path $probeOutput 'basic.xlsx') `
    (Join-Path $probeOutput 'basic-client.json') `
    --basic --timeout-seconds 180
if ($LASTEXITCODE -ne 0) { throw '基础 Excel 生命周期验收失败，请检查 harness 和结果 JSON。' }
```

## 高级 Excel 生命周期

默认高级生命周期入口由父进程创建并持有专用 Excel 实例，使用 `/x /automation /r` 只读打开已生成的夹具。父进程随后记录 PID 和启动时间、发布 owner 文件，再启动 PowerShell 子探针连接并验证该 PID；PowerShell 获取或操作 COM 对象前，必须先通过该身份检查。启动异常会写入 harness 报告；若进程句柄已取得，父进程在 `finally` 中只清理该句柄对应的实例并记录退出结果。若未取得句柄，`ownedExcelExited` 保持未知，不声称进程已退出。旧参数 `--owned-host` 保留为兼容选项，默认运行不需要传入。显式 `--probe-script` 与 `--basic` 互斥；前者仍是独立诊断模式，不参与正式生命周期验收。

在具有有效登录会话、可创建 Excel COM 实例的 Windows 环境运行：

```powershell
python ./build/PivotClientProbe/run-advanced-excel.py `
    (Join-Path $probeOutput 'advanced.xlsx') `
    (Join-Path $probeOutput 'advanced-client.json') `
    --timeout-seconds 180
if ($LASTEXITCODE -ne 0) { throw 'Excel 生命周期验收失败，请检查结果 JSON。' }
```

父进程创建自己的隐藏 Excel 实例并以只读方式打开输入；子探针在接管 COM 对象前验证进程身份，再禁用宏、链接更新和交互提示。它验证完整来源、缓存来源、字段角色、Sum 函数和 `TableRange2` 全部矩阵：首次显示应为 North 40、South 20、总计 60；将首笔 North 金额改为 15，记录刷新前的实际显示，再显式刷新，验证 45、20、65。保存到新 `.client-verified.xlsx` 副本并重开后，来源及全部汇总仍须正确。启动或身份发布失败也会生成 harness/lifecycle JSON；没有进程句柄时报告句柄未取得、退出状态未知。若生命周期工具自身在启动哨兵时失败，`lifecycle.json` 仍记录启动错误，不把清理未知记作成功。

刷新前允许完整旧结果或完整新结果，并记录 `SourceChangeVisible`；不假定所有客户端都延迟刷新。首次显示、显式刷新和重开的断言保持固定。标题可以本地化，但业务键、金额、字段角色和完整矩阵不能缺失。

首次读取若返回空来源或空矩阵，脚本使用 30 秒的被动轮询预算，每次记录 Ready、CalculationState、来源、范围和矩阵；期间不调用 Refresh、计算接口或修改来源。该预算在同步 COM 调用返回后检查，不能限制阻塞调用的墙钟时间。得到非空快照后立即执行完整断言，错误值不会靠继续轮询隐藏；预算耗尽后的空值同样失败。一次重试成功不能证明前次空快照的根因已解决，历史 Failed 结果必须保留。

脚本拒绝覆盖既有验证副本，不访问活动工作簿，不关闭用户既有 Excel。其创建的工作簿和 COM 对象在 finally 清理，清理失败不能报告 Passed。重复执行请使用新的输出目录。`RepairMode` 是传入的正常打开策略记录，不是读取了不存在的 Workbook 属性。

COM 注册或 Excel 安装信息不能替代成功结果。`0x80070520` 表示当前登录会话无法创建实例，应保留 Failed 记录，并在有效完整会话重新运行；不跳过失败断言。Excel 的成功也不能推定 WPS/LibreOffice 通过。

## 有界资源矩阵

```powershell
dotnet build/PivotClientProbe/bin/Release/net8.0/PivotClientProbe.dll --output (Join-Path $probeOutput 'rows10000-pivots4-shared.xlsx') --rows 10000 --pivots 4 --different-sources false --async true
dotnet build/PivotClientProbe/bin/Release/net8.0/PivotClientProbe.dll --output (Join-Path $probeOutput 'rows10000-pivots4-separate.xlsx') --rows 10000 --pivots 4 --different-sources true --async true
dotnet build/PivotClientProbe/bin/Release/net8.0/PivotClientProbe.dll --output (Join-Path $probeOutput 'cancel-source.xlsx') --rows 10000 --pivots 4 --different-sources false --async true --cancel-at-row 5000
```

计划矩阵为 1,000/10,000 行分别配 1 个透视、4 个同源透视、4 个不同源透视，每次用独立进程串行执行。不同源指四个真实业务 Sheet，均有完整来源和缓存检查。`cacheCount` 记录引擎实际缓存数，不要求每个透视拥有独立缓存；`enumeratedRows` 是业务序列枚举计数，不是内部来源预检或缓存扫描次数。

生成时间和分配量从输入模型构造前开始，结束于 Provider 文件导出返回；峰值工作集在随后重开验证前取样，输出大小为已提交文件大小。GC 分配计数是该进程的托管分配，峰值工作集是进程级历史峰值。独立进程启动、机器负载和 GC 会影响数据；本探针不宣称恒定内存，也没有任意性能通过阈值。

取消参数为零基来源枚举位置。取消模式先在新路径写入本工具自己的旧目标内容，随后核对同一取消令牌、旧目标字节及目录暂存清理，并记录请求到捕获的响应时间。该证据仅针对来源枚举阶段；缓存构建和序列化阶段仍由已有职责测试及后续阶段探针验证。

## 高级模式与外部硬超时

高级生命周期是默认模式，生成端需增加 `--advanced true`。高级夹具包含两行轴、两列轴、Sum/Average/Minimum/Maximum/文本 Count 和部分多选筛选；生成时断言完整来源、全部缓存记录和字段配置。`--basic` 选择前一节的基础工作簿验证脚本，不改变高级默认行为。

```powershell
dotnet build/PivotClientProbe/bin/Release/net8.0/PivotClientProbe.dll --output (Join-Path $probeOutput 'advanced.xlsx') --advanced true --async true
python ./build/PivotClientProbe/run-advanced-excel.py (Join-Path $probeOutput 'advanced.xlsx') (Join-Path $probeOutput 'advanced-client.json') --timeout-seconds 180
```

高级验证检查每个实际结果单元格与来源轴项，明确区分首次显示、修改来源后的显式刷新和保存重开。首次失败保留 Failed；诊断刷新与原生重新添加度量的结果不会改写首次状态。

当前高级夹具在本机 Excel 16.0 / Build 20430 中仍有数值验收失败：NPOI 与 ClosedXML 的文本 Count 均观察到小计或总计为空，原生新建工作簿对照也复现了文本 Count 与 Min/Max、列轴组合问题。基础/高级生成命令成功只证明 Provider 输出通过结构检查，不代表 Excel 客户端数值通过。详细结果见高级透视任务 artifacts/client，不覆盖历史产物。

Python 外围进程负责墙钟超时，记录子进程终态、专用 Excel PID 和启动时间。只有身份匹配时才清理该实例；不关闭既有 Excel。`--fault-after-owner-seconds 60 --timeout-seconds 12` 用于发布实例所有权后的阻塞故障注入，预期探针失败、timedOut=true、ownedExcelExited=true；故障注入不能视为生命周期成功。`verify-owned-host-lifecycle.py` 不传 `--owned-host`，验证默认入口在附着前和附着后故障时只清理本次启动的 Excel。该工具的“既有实例”哨兵也是本任务创建的独立进程，不代表用户的交互实例或其他登录会话；因此证据范围不包含用户现有 Excel 的全局生命周期。

```powershell
python ./build/PivotClientProbe/verify-owned-host-lifecycle.py `
    (Join-Path $probeOutput 'advanced.xlsx') `
    (Join-Path $probeOutput 'owned-host-lifecycle')
if ($LASTEXITCODE -ne 0) { throw '专用 Excel 生命周期故障注入未满足预期。' }
```

`--probe-script` 可运行显式指定的本地诊断脚本，结果须声明 DiagnosticOnly；不以诊断完成替代上述全矩阵验收。已有结果路径及工作簿不覆盖。

## 数字格式入口诊断

使用 `--number-format-probe true` 生成单独的诊断工作簿，对比引擎公开 SetFormat、Format 属性、编号及数据区域样式入口，保存调用后的实际格式、重开格式和相关 OOXML。同名 `.probe.json` 标记 diagnosticOnly，不代替 Provider 职责或客户端数值验收。

```powershell
dotnet build/PivotClientProbe/bin/Release/net8.0/PivotClientProbe.dll --output (Join-Path $probeOutput 'number-format.xlsx') --number-format-probe true
```

当前 0.105.1 的 SetFormat 入口识别内建格式，但自定义 `0.0` 调用后文本为空；Format 属性保留自定义文本。Provider 已保留内建编号识别并对自定义格式使用属性入口，双 TFM 的内建/自定义格式和组合职责已通过；这不解决高级文本 Count 的客户端问题。

## 外围工具故障职责测试

不安装或启动 Office 时，可先运行外围进程管理测试：

```powershell
python ./build/PivotClientProbe/tests/test_owned_runner_failure_contracts.py -v
if ($LASTEXITCODE -ne 0) { throw '外围工具故障职责验证失败。' }
```

这些测试调用真实 runner，使用本任务创建的 Python 子进程替代 Excel 和 COM 子探针，覆盖启动失败、身份查询失败、owner 写入失败、子进程启动失败、非零退出、超时、成功终态及非法/非对象 JSON，并覆盖基础模式的成功、超时和子进程失败清理。测试通过真实进程句柄确认本次进程已退出，同时另一个任务哨兵仍存活；结束时清理哨兵。它们证明外围工具的控制流、文件报告和进程清理合同，不证明 Excel 附着、工作簿统计结果或用户交互实例行为。

子探针结果无法读取、不是合法 JSON 或不是 JSON 对象时，runner 保留清理结果，写出 `terminalError` 和 Failed harness，并返回非零退出码；不会因为子进程退出 0 而报告通过。正常 Office 的连接前后超时与完整首开/刷新/重开仍须执行上面的客户端验证。
