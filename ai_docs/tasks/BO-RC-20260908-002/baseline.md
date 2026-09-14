# BO-RC-20260908-002 基线记录

- Task-ID：`BO-RC-20260908-002`
- Branch：`master`
- HEAD：`e4093f554e7e81d642ff6722a1f8e44d185d39da`
- 执行器：Codex；工作树未执行 `git add`、`git commit`、`git push`、PR 或发布。
- 环境：Windows 10 `10.0.19045.6466`，x64，.NET SDK `10.0.400`，运行时 `.NET 6.0.36` / `.NET 8.0.30`。

## 依赖基线

- 初次 `dotnet restore Bing.Offices.sln --locked-mode` 暴露了已有 lockfile 漂移（`NU1004`）；随后用 `--force-evaluate` 重新生成锁文件。
- 重新执行 `dotnet restore Bing.Offices.sln --locked-mode --verbosity minimal`：通过。
- 两个 PackageReference-only consumer 使用独立缓存和本地最终包 feed；最终 locked restore：通过。net6 仅有 `NETSDK1138` 生命周期警告。
- `.gitignore` 不再忽略 `packages.lock.json`；solution 参与项目的 lockfile 作为可审计变更保留。

## 构建与资产

- `dotnet build Bing.Offices.sln -c Release --no-restore /m:1`：通过，0 warning / 0 error。
- `dotnet build tests/Bing.Offices.ResourceProbe/Bing.Offices.ResourceProbe.csproj -c Release --no-restore /m:1`：通过，0 warning / 0 error。
- 最终生产包位于 `artifacts/packages/`，包含 `Bing.Offices.Abstractions`、`Bing.Offices.Core`、`Bing.Offices.Npoi` 2.0.0；NPOI 同时包含 `lib/net6.0` 和 `lib/net8.0`，Abstractions/Core 包含 `lib/netstandard2.0` 及 XML 文档。

## 初始风险与本轮处理

- CI 的 locked restore 依赖 lockfile；本轮修复可追踪 lockfile 策略并把 package consumer 接入 CI。
- ResourceProbe 初始只创建 `Math.Min(requested, 1)` 个请求；本轮改为创建完整请求数并用实际队列事件统计。
- Excel 模板 Async 保持 NPOI 同步读取边界；通过双 TFM characterization 和文档明确化，不新增模板 staging。
