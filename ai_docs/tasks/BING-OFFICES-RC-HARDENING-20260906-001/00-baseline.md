# 基线报告

## 任务信息

- Task ID：`BING-OFFICES-RC-HARDENING-20260906-001`
- 基线分支：`master`
- 基线提交：`958e5b4886c2fe8df80ece3218d9fab1c57a0ec6`
- 采集时间：2026-09-06（Asia/Shanghai）
- 执行器：Codex
- 实现审查文件：`ai_docs/codebase-analysis/bing-offices-implementation-review-20260906.md` 不存在，未伪造其结论。

## 工作区

开始执行时 `git status --short` 仅包含本任务计划目录；未发现生产代码、测试或项目文件的既有未提交修改。

## 环境

- OS：Windows 10.0.19045，RID `win-x64`
- SDK：10.0.400，MSBuild 18.9.6
- 已安装 SDK：3.1.426、6.0.428、8.0.424、10.0.300、10.0.400
- 已安装运行时：netcoreapp3.1、net6.0、net8.0、net9.0、net10.0
- `global.json`：不存在

## 项目目标基线

- `Bing.Offices.Abstractions`、`Bing.Offices.Core`：`netstandard2.0`
- `Bing.Offices.Npoi`：`framework.props` 当前为 `net8.0;net6.0;netcoreapp3.1`
- Unit：`net8.0;net6.0;netcoreapp3.1`
- Integration：`net6.0;net8.0`
- 现有 Release 输出目录：`output/release/netstandard2.0`、`netcoreapp3.1`、`net6.0`、`net8.0`

## 基线命令与结果

| 命令 | 结果 | 说明 |
| --- | --- | --- |
| `dotnet --info` | PASS | 环境信息如上 |
| `dotnet restore Bing.Offices.sln --locked-mode` | BLOCKED | 多个 lock 文件与当前项目 PackageReference 不一致（NU1004）；ResourceProbe、ApiSnapshot、Benchmark 和 BuildScript 另受 NuGet TLS/凭证错误（NU1301）影响 |
| `dotnet build build/ApiSnapshot/ApiSnapshot.csproj --no-restore -c Release` | BLOCKED | 现有 assets 仍尝试解析 `System.Reflection.MetadataLoadContext`，NuGet 源 TLS 失败 |

## 已知基线差距

1. 执行阶段已从基线 commit 的临时副本生成 `build/api-snapshot-baseline.json` 和 before snapshot；审批字段仍为空，不能解除 API 门禁。
2. `ExcelStreamExtensions`/`CsvStreamExtensions` 在 exporter 公共观察边界外提交文件。
3. JSON/XML mapping loader 保留空诊断输出重载和 public `ExcelMappingDiagnostic`。
4. API canonicalizer 未编码修饰符、约束、参数名/默认值、访问器可见性和关键 attribute。
5. `TryAddPicture` 在 `AddPicture` 后失败时可能返回 `false` 并留下副作用。
6. NPOI 与测试仍包含 EOL TFM。

## 基线冻结说明

由于 restore/ApiSnapshot 工具受外部 NuGet TLS 阻断，成员级 before snapshot 将在本地可用的既有 Release 输出基础上生成；所有无法取得的环境证据在最终报告中标记为 BLOCKED，不以猜测或自动批准替代。
