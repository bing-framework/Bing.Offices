# 基线报告

- 任务：`BING-OFFICES-RELEASE-HARDENING-20260831-001`
- 状态：`IN_PROGRESS`
- 基线时间：2026-08-31
- 工作区：Windows 10 x64，.NET SDK 10.0.300；已安装并用于强制回归的运行时为 .NET 6.0.36、8.0.27。
- 用户提供的 `ai_docs/codebase-analysis/bing-offices-implementation-review-20260831.md` 与 `merged_20260831085402.md` 当前未找到，标记为 `INPUT_MISSING`；用户提供的 SHA 暂无对应文件可核验。

## 工作区

- 执行前源码没有新的已跟踪差异；任务目录为本轮新建未跟踪内容。
- 禁止执行 `git add`、`git commit`、`git push`、Tag、PR、NuGet publish、reset、clean。
- 前次会话提示的四个文件在实施前已重新读取并遵循现状。

## 当前验证

- `dotnet restore Bing.Offices.sln --locked-mode`：失败，多个项目报告 `NU1004`，锁文件记录与项目 PackageReference 表达式不一致；未修改锁文件。
- 后续 build/test/pack/consumer/benchmark 结果在对应报告中记录，并绑定当前工作树。

## 初始判断

功能主链已存在；当前发布成熟度暂定 No-Go。主要执行重点为 selector 单次解析、故障矩阵复验、API 分类收敛、包消费可复现性、Benchmark 原始证据和发布材料。
