<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-MINIEXCEL-PROVIDER-20260915-001
AI_EXECUTION_FINISHED_AT: 2026-09-16T01:44:59+08:00

# 实施执行报告

## 执行结论

计划内 MiniExcel Provider 核心范围已实现并完成必要构建、测试、包和双 TFM API 验证。长耗时 BenchmarkDotNet 自动生成项目受 NuGet SSL/凭证环境阻断，500K/1M 未执行；这两项作为明确限制记录，不作为通过声明。

## 任务信息

- Task ID：`BING-OFFICES-MINIEXCEL-PROVIDER-20260915-001`
- 执行器：Codex
- 目标：新增独立 `Bing.Offices.MiniExcel`，复用 Provider-neutral Request/Mapping/Converter/Validation/Exception/File Committer 契约。
- 版本：MiniExcel `1.46.0`；Provider 目标框架 net6.0、net8.0。
- 执行开始：2026-09-15T14:55:31.076Z；执行完成：2026-09-16T01:44:59+08:00。
- Git 操作边界：未自动 commit、未自动 push、未创建 PR。

## 计划执行情况

| Phase | 状态 | 说明 |
|---|---|---|
| 0 Baseline/API Spike | 完成 | 锁定 MiniExcel 1.46.0，完成真实 SaveAs/Query/async/资源行为验证 |
| 1 Contract/共享预检 | 完成 | XML 注释中性化；Core ZIP/XML preflight；NPOI 回归 |
| 2-8 Core/Exporter/Importer | 完成 | XLSX、多 Sheet、映射、动态列、转换、校验、关系、取消和原子提交 |
| 9 Unsupported/DI | 完成 | unsupported preflight、DI 扩展、注册顺序语义文档化 |
| 10 API/Package Consumer | 完成 | 双 TFM API snapshot、4 包 pack、net6/net8 PackageReference consumer |
| 11-13 Unit/Integration/Contract | 完成 | MiniExcel 职责测试、真实文件集成、NPOI 共享契约回归 |
| 14 Benchmark/Resource | 部分完成 | 100K controlled probe 和资源矩阵完成；BDN/500K/1M 受环境限制 |
| 15 Documentation | 完成 | README、Excel provider 文档和任务报告 |
| 16 Release readiness | 完成 | CI pack/verify、全量测试、API compare、diff check |

## 已完成事项

- 新增 `src/Bing.Offices.MiniExcel`，公开 exporter、importer 和 DI extension，不暴露 MiniExcel/NPOI 类型。
- 新增 Core `ExcelXlsxZipPreflight`，NPOI 和 MiniExcel 使用相同 ZIP/XML 安全预算、DTD/路径/压缩比和取消边界。
- MiniExcel 导出使用真实 `SaveAs`/`SaveAsAsync`，导入使用真实 `Query`/`QueryAsync`，映射统一走 Core plan。
- 实现普通 XLSX、多 Sheet、固定/动态列、ValueMap、Converter、Validation、Unique、Relations、Stream/File/byte[] 和原子文件提交。
- 对 XLS、模板、样式/数字格式/列宽/布局、multi-header、图片、批注、Chart、Failure Workbook、元数据和隐藏 Sheet 在写入/解析前 fail-fast。
- 接入 solution、CI pack、API snapshot baseline、双 TFM PackageReference consumer、Benchmark list/probe 和 Provider 文档。
- 新增真实文件 MiniExcel integration test；测试结果写入 `integration-test-report.md`。

## 部分/未完成事项

- BenchmarkDotNet Dry 找到 6 个 MiniExcel benchmark，但自动生成项目因 `NU1301`（api.nuget.org SSL/凭证）无法构建，0 次有效迭代。
- 500K/1M 长测未执行；当前性能证据只覆盖 100K controlled probe。
- MiniExcel 导入先缓冲到 MemoryStream，未承诺完整端到端 streaming importer；已在资源/性能报告中标注容量风险。
- API baseline 的 `approvedBy=task-plan-additive-api` 是任务内 additive 记录，不等同人工成员审批。

## 修改文件

- 生产：`src/Bing.Offices.MiniExcel/**`、`src/Bing.Offices.Core/Bing/Offices/IO/ExcelXlsxZipPreflight.cs`、Core/NPOI friend/preflight、Abstractions XML 注释。
- 工程：`Bing.Offices.sln`、`version.dev.props`、CI、API snapshot/build、双消费者、Benchmark 项目。
- 测试：MiniExcel unit test、MiniExcel integration test、测试项目引用和 Public API contract。
- 文档：README、`docs/excel/09-providers.md`、本任务矩阵与报告。

## API/数据/配置变化

- 新增 fourth package `Bing.Offices.MiniExcel`，只做 additive public surface。
- 新增 `MiniExcelPackageVersion=[1.46.0]`；不变更已有公共成员，不删除 NPOI API。
- Core 新增 internal ZIP preflight helper；NPOI 行为通过 position reset 回归测试保持不变。
- CI 包数量由 3 增至 4；MiniExcel nuspec 无 NPOI dependency/tag。

## 测试结果

- MiniExcel filtered Unit：10/10（net6.0、net8.0）。
- Full Unit：739/739（net6.0、net8.0）。
- Full Integration：39/39（net6.0、net8.0）。
- NPOI shared XLSX preflight：27/27（net6.0、net8.0）。
- Docs：10/10（net8.0）。
- Package consumers：.NET 6.0.36 和 .NET 8.0.30 均输出 `package-consumer-ok`。
- Controlled probe：100,000 行 sync/async 均完成并读回，原始 JSON 为 `artifacts/benchmark-miniexcel-probe-100k-final.json`。
- ResourceProbe：`scenarios=36 status=passed approval=BLOCKED`；36 个场景均通过，未提供历史资源任务审批参数，故仅保留工具返回的审批状态。

## Build/Typecheck/Lint/Format

- `dotnet build .\Bing.Offices.sln -c Release --no-restore --nologo`：0 warning、0 error。
- 四个包 pack 成功：4 `.nupkg`、4 `.snupkg`；MiniExcel 包内容和 nuspec 已检查。
- API snapshot capture/compare：net6.0、net8.0 均通过。
- `git diff --check`：无 whitespace error；LF/CRLF 提示是 Git working-tree line-ending warning。
- `rg` 审计：MiniExcel production/Core preflight 无 `Task.Run`、`.Result`、`.Wait()`；MiniExcel 源码无 NPOI 引用。

## 计划偏差

1. 原计划将 MiniExcel 集成测试作为候选独立项目；实际复用现有 Integration 项目并增加 MiniExcel ProjectReference，减少测试工程数量，仍保持真实路径和双 TFM 隔离。
2. 原计划考虑抽取更多 NPOI helper；实际只抽取被两 Provider 证明复用的 ZIP/XML preflight，关系和行语义继续复用既有 Core contract，避免无证据的大范围重构。
3. MiniExcel 未证明完整样式/模板/图片/批注语义，因此按计划标为 unsupported 并实现 fail-fast，而不是静默降级或宣称支持。

## 基线问题

- 初始 sandbox restore/BenchmarkDotNet 自动项目遇到 NuGet SSL/凭证错误；使用批准的 escalated restore 完成项目依赖，BenchmarkDotNet 长测仍保留为受限。
- API baseline 采用 additive candidate identity，并明确写入任务内审批标记；发布前需要维护者确认。
- net6 构建出现 SDK 的 EOL `NETSDK1138` warning，仅来自包消费者目标框架，不影响编译或测试结果。

## 已知问题

- 100K probe 的分配量约 1.28 GB，表明大数据部署需要容量评估；不能据此承诺低内存。
- 导入缓冲整个输入流，超出 `MaxInputBytes` 会在 parser 前失败；这是一项已记录的设计边界。
- 默认 DI 保持先注册 Provider 生效；文档要求一个容器选择一个 Provider，resolver 不在本任务范围内。

## 风险与回归关注点

- MiniExcel 升级到 2.x 前必须重新验证 API/TFM/async/取消和包依赖，不能直接替换版本。
- unsupported preflight 的新增检查应继续保持“零字节写入”和稳定 provider/operation/stage。
- Core ZIP preflight 变更会影响 NPOI 与 MiniExcel，任何预算调整都要重跑双 TFM NPOI 预检矩阵。
- 高并发和 500K/1M 需在网络、磁盘和 NuGet 可用的基准机补跑，关注 working set、GC、临时文件和句柄。

## Reviewer 注意事项

- 重点检查 `MiniExcelExcelExporter.ValidateRequest/ValidatePlanCapabilities` 是否覆盖所有被忽略的高级配置，并确认 unsupported 不会在写入后才暴露。
- 重点检查 `ExcelXlsxZipPreflight` 的 stream position、DTD、路径穿越、重复 entry 和资源预算是否保持 NPOI 原有语义。
- 重点检查 API baseline 的 additive approval marker 是否在合并前获得仓库维护者确认。

## Git 状态

- 工作树包含本任务生产、测试、文档和基线变更；未覆盖或还原用户已有改动。
- 未自动执行 `git add`、`git commit`、`git push` 或 PR 创建。
