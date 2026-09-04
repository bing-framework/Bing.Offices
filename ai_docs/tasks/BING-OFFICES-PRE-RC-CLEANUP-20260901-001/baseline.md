# Phase 0 基线

## 任务信息

- Task ID：`BING-OFFICES-PRE-RC-CLEANUP-20260901-001`
- 基线时间：2026-09-01
- 分支：`master`
- HEAD：`54c641ded99e4b147aa81bb325c36021382af46b`
- OS：Windows 10.0.19045，x64
- SDK：.NET SDK 10.0.300，MSBuild 18.6.3
- 已安装运行时：.NET 6.0.36、8.0.27、9.0.16、10.0.8；未安装 netcoreapp3.1/net5/net7 运行时
- 解决方案：`Bing.Offices.sln`
- NuGet 源：nuget.org、Microsoft Visual Studio Offline Packages
- 锁定：项目和测试公共 props 启用 `RestorePackagesWithLockFile`

## 工作树

初始 `git status --short` 仅显示：

```text
?? ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/
```

初始 `git diff --stat` 无输出；`git diff --check` 无输出。当前新增内容属于本任务证据目录，不回滚、不覆盖。

## 输入核验

用户要求的 `ai_docs/codebase-analysis/` 评审与方法论文件在当前工作区不存在；仅按当前源码、项目配置和运行结果取证。前序 `BING-OFFICES-RELEASE-HARDENING-20260831-001/plan.md` 存在，但其历史结论不替代本任务运行证据。

## 构建/测试入口

`dotnet sln Bing.Offices.sln list` 确认包含 3 个生产项目、3 个测试项目、ProfileFixtures、ResourceProbe、ApiSnapshot 和 Benchmarks。生产依赖方向为 `Abstractions <- Core <- Npoi`；Unit 为 `tests/Bing.Offices.Tests`，Integration 为 `tests/Bing.Offices.Tests.Integration`，Docs 为 `tests/Bing.Offices.Docs.Tests`。

## 当前静态阻断

- `ExcelMappingDocumentFactory.Create` 接收 `requestConfiguration` 但当前返回值未合并它。
- v1 JSON/XML 迁移的非目标方向当前构造空配置而非 `null`。
- `NpoiRelationBinder.Bind` 通过反射 `Invoke`，未在该入口解包 `TargetInvocationException`。
- CSV `ProtectFormula` 当前仅检查第一个字符。
- `CsvImportOptions.Validate` 当前未校验 `MaxTrackedUniqueValues` 与 `UniqueComparison`。
- ResourceProbe 当前先独立创建 Workbook 再执行导入。
- `AddNpoi` 当前返回 `void`；存在 compatibility/execution-detail public API 和生产 `NotImplementedException` 候选。

## 基线状态

P0-01：`VERIFIED`（环境、项目、Git 状态和缺失输入已记录）。
P0-02：`IN_PROGRESS`，等待 Release build/test/pack 运行证据。
P0-03：`IN_PROGRESS`，待完成全仓候选符号清单。
