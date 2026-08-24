# TASK-BING-OFFICES-20260821-MAPVAL-V2 基线

## Git

- Task ID：`TASK-BING-OFFICES-20260821-MAPVAL-V2`
- 分支：`master`
- HEAD：`73883f709f8eb9e58cd948db0bf90e82ca44a661`
- 状态：`master...origin/master`
- Task 启动后新增：`.agents/runtime/current-task.json`、本 Task 目录。
- Task 启动前已有用户改动：13 个 `Metadata/Excels` 源文件删除，未恢复、未覆盖。

## 包与 Target Framework

- `Bing.Offices.Abstractions`：`netstandard2.0`，PackageId 保持不变。
- `Bing.Offices.Core`：`netstandard2.0`，引用 Abstractions、Bing.Utils `1.5.0`、CsvHelper `33.1.0`。
- `Bing.Offices.Npoi`：`net8.0;net7.0;net6.0;netcoreapp3.1`，引用 Core、NPOI `[2.7.4]`。
- 版本：`1.0.0`，来自 `version.props`。
- Unit：`net8.0;net7.0;net6.0;net5.0;netcoreapp3.1`。
- Integration：`net6.0;net8.0`。
- Docs Consumer：`net8.0`。
- Benchmark：`net8.0`，BenchmarkDotNet `0.14.0`。

## Compatibility Mode

判定为 `MIGRATION_CURRENT_MAJOR`：当前版本为 `1.0.0`，用户要求保留三个已发布包及旧 API 兼容迁移，且没有 next-major 发布信号。后续 API 采用 additive + Obsolete forwarding；不改变已发布 `AddNpoi(): void` 返回类型。

## 验证基线

### Restore

命令：`dotnet restore Bing.Offices.sln --locked-mode`

结果：`FAILED`，退出码 1，耗时约 7.6 秒。

主要错误：`NU1004`。当前项目声明使用 `>=` 依赖范围，而多个现有 lock 文件保存为带方括号的范围，例如 `Bing.Utils:[1.5.0, )`、`NPOI:[2.7.4, 2.7.4]`；NuGet 判定 lock 文件与项目依赖不一致，建议 `--force-evaluate`。本次尚未修改 lock 文件。

### Build/Test/Pack/Benchmark

- `dotnet build Bing.Offices.sln -c Release --no-restore`：`PASSED`，73 个既有警告，0 错误。
- Unit net6：`PASSED`，135/135。
- Unit net8：`PASSED`，135/135。
- Integration net6：`PASSED`，10/10。
- Integration net8：`PASSED`，10/10。
- Docs Consumer net8：`PASSED`，2/2。
- `dotnet pack`：`TODO`，在 API 稳定后执行。
- Pack Consumer：`TODO`。
- Benchmark：`TODO`。

## 设计文档基线

- `ai_docs/Bing.Offices-validation-mapping-refactor-solution-v2-20260821.md`：未找到。
- `Bing.Offices-implementation-review-20260821.md`：未找到同名文件或变体。
- 可用证据：根 `AGENTS.md`、`.github/copilot-instructions.md`、`ai_docs/excel/*`、源码、测试、CI。