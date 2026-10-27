# 本地候选包与新 API 消费验证

运行统一脚本可验证当前源码打出的 NuGet 包，而不需要修改批准的 API 基线或发布包。需要 PowerShell 7、.NET 6/8 SDK 和 Git，首次恢复第三方依赖需要访问 NuGet。

在仓库根目录执行，输出目录必须尚不存在：

```powershell
pwsh -File build/Verify-PackageCandidate.ps1 -OutputDirectory TestResults/package-candidate-local
```

脚本依次完成以下检查，任何失败均以非零退出并记录日志：

1. 构建并打包 Abstractions、Core、Npoi，验证三个 nupkg、三个 snupkg 及库/XML 文档。包版本从实际 nuspec 读取，不覆盖项目版本。
2. 运行既有身份契约自检，捕获本轮源码、程序集和包的逻辑身份。
3. 将消费者复制到仓库外临时目录，从空缓存恢复。Bing.Offices.* 只能从本轮本地源取得；项目引用、同版本全局缓存和仓库构建导入均不参与消费验证。
4. 运行 net6/net8 的 NPOI 消费者以及两个框架的 Core-only CSV 消费者。除旧 Workbook 入口外，还检查单表入口、完整数据/错误和状态摘要；Core-only 的恢复图不得包含 NPOI。
5. 将当前程序集的 API 与批准基线比较，并对照捕获清单检查当前源码/产物身份。缺失清单、损坏身份、API 不匹配必须被拒绝。

成功标记为 `package-candidate-ok consumers=4 negativeChecks=3`。`summary.json` 给出本轮状态，`logs/` 保存各命令输出，`consumers/` 保存依赖恢复图，`capture/` 与 `comparison/` 保存双 TFM 快照、候选和 API 差异；`negative-cases/` 中的文件仅用于故意失败的测试，不是批准基线。消费者及包缓存退出时清理，打包产物和证据保留。

## API 合同与候选身份

`build/api-snapshot-baseline.json` 是批准的公共 API 合同。正常代码修复后，新候选身份不应写回此文件来使检查通过。先捕获，再显式传入当前清单：

```powershell
dotnet run --project build/ApiSnapshot/ApiSnapshot.csproj -c Release -- --root output/release --packages <包目录> --repository . --capture true --output <捕获目录>
dotnet run --project build/ApiSnapshot/ApiSnapshot.csproj -c Release -- --root output/release --packages <包目录> --repository . --baseline build/api-snapshot-baseline.json --candidate <捕获目录>/api-candidate.json --output <比较目录>
```

`--candidate` 只提供本轮身份；成员与审批仍由 `--baseline` 决定。未传 `--candidate` 时保留旧行为，使用基线自身记录的身份。`--capture true` 不能与 `--candidate` 同时使用。捕获成功不是 API 审批，也不表示比较通过。

身份沿用现有 `logical-v3` 规则，规范化 API、源码与包内容，用于跨环境比较；它不是原始 DLL/nupkg 的逐字节签名或源码构建证明。脚本在同一轮重新构建、打包、消费，再比较，避免复用历史产物冒充当前候选。

## CI 与验证范围

`.github/workflows/ci.yml` 的 `Verify package candidate` 使用同一脚本，成功或失败均尝试上传证据。本地 Windows 通过不能代替 GitHub Ubuntu run；远端结果需关联实际运行提交和 artifact。

本流程验证包接入、API 合同和候选一致性，不执行发布、生产部署、重型容量或性能门禁。原始输入 128 MiB 上限不代表 Office DOM 峰值内存限制。生产资源/并发证据应继续使用对应部署门禁。
