# Phase 0 基线

## 身份

- 采集时间：2026-09-06T11:25:39+08:00
- OS：Windows 10.0.19045，win-x64
- Branch：`master`（跟踪 `origin/master`）
- Commit：`9d78ab76e28891ac9b7e0e558f4df669721a4ea6`
- 执行前 dirty state：clean
- SDK：3.1.426、6.0.428、8.0.424、10.0.300、10.0.400；实际默认 10.0.400
- Runtime：.NET Core 3.1.32、.NET 6.0.36、.NET 8.0.27/8.0.30 等
- 生产 TFM：Abstractions/Core netstandard2.0；NPOI netcoreapp3.1/net6.0/net8.0

## Restore 与 Build

`dotnet restore Bing.Offices.sln --nologo`：沙箱内首次因 NuGet TLS/凭据及缺失本地 source 失败；受控网络权限重跑成功，11 个项目均恢复。

`dotnet build Bing.Offices.sln -c Release --no-restore --nologo`：成功，0 errors / 28 warnings。警告包括 netcoreapp3.1 依赖不再受支持、API snapshot nullable 上下文、过时 NPOI API 和 xUnit analyzer。

## 测试基线

| 项目/TFM | Passed | Failed | Skipped | 结论 |
| --- | ---: | ---: | ---: | --- |
| Unit netcoreapp3.1 | 453 | 1 | 0 | Failed：API hash |
| Unit net6.0 | 453 | 1 | 0 | Failed：API hash |
| Unit net8.0 | 453 | 1 | 0 | Failed：API hash |
| Integration net6.0 | 15 | 0 | 0 | Passed |
| Integration net8.0 | 15 | 0 | 0 | Passed |
| Docs net8.0 | 11 | 0 | 0 | Passed |

唯一 Unit 失败：`PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`，Abstractions expected `7F9A...`、actual `1A080...`。这是修改前基线失败，不允许直接更新 hash 隐藏差异。

原始结果：`artifacts/test-results/unit`、`artifacts/test-results/integration`、`artifacts/test-results/docs`。

## Public surface 基线

- 当前源码分类总计 180 个导出类型：User API 82、Provider User API 6、Provider SPI 9、Execution detail 83、Compatibility 0。
- NPOI 七个要求保留的扩展容器均为 public；分类中 DI 扩展被记为 User API，其余六个为 Provider User API。
- 独立 API snapshot CLI 无法完成：代码默认读取 `build/api-snapshot-baseline.json`，该文件在仓库不存在。Unit 内嵌 hash 同样显示正式 baseline 漂移。

## 输入证据缺口

- `ai_docs/codebase-analysis/bing-offices-implementation-review-20260905.md` 不存在。
- 未发现 Mustang 方法论路由器或 manifest，不伪造 framework-source 运行结果。
- 当前仓库未发现 API snapshot baseline JSON。

## 基线结论

状态：`No-Go`。源码和主功能测试可运行，但正式 API baseline 失败，附件列出的异常误分类、弃用表面、资源与性能证据仍未闭环。

