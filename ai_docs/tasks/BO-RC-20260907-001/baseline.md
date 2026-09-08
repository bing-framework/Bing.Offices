# BO-RC-20260907-001 基线

## 环境

- Branch/commit：`master` / `4ecfe3fa7acea6b9d22d00e52e102c5ed1cf31bc`
- OS：Windows 10 x64
- SDK：6.0.428、8.0.424、10.0.400；执行默认 SDK 为 10.0.400
- Runtime：.NET 6.0.36、8.0.30 可用
- 初始工作区：clean；执行后仅包含本任务实现与证据

## 还原

首次 `dotnet restore Bing.Offices.sln --locked-mode` 暴露既有 lockfile 漂移（NU1004），ApiSnapshot/BuildScript 同时受到 NuGet TLS/凭证限制（NU1301）。使用受控网络和 `--force-evaluate` 后 solution restore 成功，锁文件更新为双 TFM 依赖图。

## 当前编译/测试基线

| 项目/TFM | 结果 |
| --- | --- |
| Abstractions/Core `netstandard2.0` | Release build 通过 |
| NPOI `net6.0` | Release build 0 warning/0 error |
| NPOI `net8.0` | Release build 0 warning/0 error |
| Unit `net6.0` | 495 passed，1 API baseline approval blocked |
| Unit `net8.0` | 495 passed，1 API baseline approval blocked |
| Integration `net6.0` | 15/15 passed |
| Integration `net8.0` | 15/15 passed |
| Docs `net8.0` | 10/10 passed |
| Async direct tests | net6/net8 均通过；当前 4 个场景 |

Unit 唯一失败是 `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot` 的维护者审批字段为空，不是代码编译/行为失败；不得由 Executor 自动填写批准人。

## 发布包

任务专属 `artifacts/packages` 已重新使用不带 `--no-build` 的 `dotnet pack` 生成三个生产包；NPOI 包包含 `lib/net6.0` 与 `lib/net8.0`。旧 `--no-build` 包曾复用过期 bin 输出，已丢弃其结论。

## 基线限制

- .NET 6 已 EOL，SDK 产生 `NETSDK1138`；仍按用户要求作为兼容 TFM 验证，并在最终报告披露风险。
- API baseline 仍只有已批准结构的历史 net8 记录，`approvedBy/approvedAt` 为空；新 Async/public Provider/namespace 差异必须人工审批。
- macOS/Linux 本地 runner 未执行；CI 已更新为安装并测试 net6/net8。
