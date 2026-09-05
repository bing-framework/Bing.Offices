# RC Hardening 基线

## 任务

- Task ID：`BING-OFFICES-RC-HARDENING-20260904-001`
- 基线时间：2026-09-04
- 分支：`master`
- HEAD：`698ed34dd6055269760cb526b76a7936fb70f9dd`
- 工作树：源码、测试、配置无已跟踪改动；仅本任务目录为未跟踪输入。
- 上一任务 `BING-OFFICES-PRE-RC-CLEANUP-20260901-001` 的工作树改动仍属于当前共享工作区，不回滚、不覆盖。

## 环境

- OS：Windows 10.0.19045，x64
- SDK：.NET SDK 10.0.400
- MSBuild：18.9.6
- 已安装 runtime：netcoreapp3.1、net6.0、net8.0 及其它版本
- 解决方案：`Bing.Offices.sln`
- 生产依赖方向：Abstractions `netstandard2.0` <- Core `netstandard2.0` <- NPOI `netcoreapp3.1/net6.0/net8.0`
- 测试框架：xUnit + Microsoft.NET.Test.Sdk
- Benchmark：BenchmarkDotNet 0.14.0
- 主要包：NPOI 2.7.4、CsvHelper 33.1.0
- 未发现仓库级 `NuGet.Config`、Mustang methodology manifest 或 `ai_docs/codebase-analysis/bing-offices-implementation-review-20260904.md`。

## 真实项目

由 `dotnet sln .\\Bing.Offices.sln list` 确认：三个生产项目、Unit、Integration、Docs、ProfileFixtures、Benchmark、ApiSnapshot 和 BuildScript。

## 基线规则

- 现有共享工作区差异视为用户/前序任务输入，后续只在本任务范围内最小修改。
- 本任务不执行 `git add`、commit、push、tag、publish、reset 或 clean。
- 所有任务报告与源码按 UTF-8 处理。
- 首次构建和测试失败必须原样记录，不通过删除断言、跳过测试或修改基线掩盖。
