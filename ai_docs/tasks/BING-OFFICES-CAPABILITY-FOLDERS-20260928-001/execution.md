<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-CAPABILITY-FOLDERS-20260928-001
AI_EXECUTION_FINISHED_AT: 2026-09-28T15:07:49Z

# 实施执行报告

## 执行结论

已按批准计划完成能力目录归类。108 个 Abstractions 类型文件中，106 个移入子目录，`IExcelExporter` 与 `IExcelImporter` 保留在对应目录根部。所有公开类型仍使用原命名空间和签名；NPOI 实体布局及 Abstractions 的内部辅助类型迁入对应能力命名空间。未修改业务逻辑。

## 任务信息

- 计划：`plan.md`
- 执行器：Codex
- 状态：`COMPLETED`

## 计划执行情况

| 项目 | 结果 |
|---|---|
| Entities 归类 | 36 个文件按 Layout、Lists、Calculated、Dynamic、Footers、Subtotals、Import、Export 归类 |
| Exports 归类 | 40 个文件移入 Workbook、Sheets、DynamicColumns、Reports、Charts、Content、Streaming；通用接口留在根部 |
| Imports 归类 | 30 个文件移入 Workbook、Sheets、Batch、Failures、Validation、Resources；通用接口留在根部 |
| Provider 对齐 | 3 个 NPOI 实体布局内部类型使用 `Bing.Offices.Npoi.Entities`，两个调用文件更新引用 |
| 文件与文档引用 | 项目文件无相关显式源码 Include；现行文档无需要更新的路径；历史任务记录保留 |

## 已完成事项

- 逐文件对照 HEAD，106 个移动的 Abstractions 文件类型主体保持一致；5 个 NPOI 文件仅更改命名空间和 `using`。两处动态列辅助类型调用只更新了类型引用。
- 移动的公开类型仍分别属于 `Bing.Offices.Entities`、`Bing.Offices.Exports`、`Bing.Offices.Imports`。
- `ExcelImportPolicies.cs` 随迁移改名为其唯一类型对应的 `ExcelImportFailureOptions.cs`。
- 保留原有中文 XML 注释、UTF-8 BOM、LF 与文件末尾换行。

## 部分/未完成事项

无。

## 修改文件

- 生产源码：Abstractions 的 Entities、Exports、Imports 目录内 106 次文件移动及必要引用；NPOI 的 3 个实体执行类型与 2 个调用文件。
- 任务记录：本目录的 `plan.md`、`execution.md` 和 `artifacts/api/` 候选快照。

## API/数据/配置变化

公开 API、数据格式、运行配置、包版本均无变化。内部类型命名空间按能力调整。双 TFM 的 16 个程序集/TFM API 成员集合与批准基线逐项比较，新增 0、删除 0；未更新基线。

## 测试结果

- `dotnet test Bing.Offices.sln -c Release --no-build --no-restore --filter 'Category!=Large' -m:1 --verbosity quiet`：全部通过，共 3,374 项；NPOI 的 net6.0/net8.0 各 718 项通过。
- API 快照工具对 net6.0、net8.0 采集成功，候选结果保存在 `artifacts/api/`；成员集合与 `build/api-snapshot-baseline.json` 比较为零差异。此处验证的是成员差异，未将候选身份或包身份作为本次发布审批。

## Build/Typecheck/Lint/Format

- Abstractions 与 NPOI 职责级 Release 构建：0 警告、0 错误。
- `Bing.Offices.sln` Release 构建：0 错误；4 个既有的 Microsoft.Bcl.Memory 对 net6.0 的支持警告。
- 111 个改动 C# 文件严格 UTF-8 BOM、LF、单个末尾换行检查通过；`git diff --check` 通过。

## 计划偏差

无行为或公开 API 偏差。为保持“一个文件一个非 partial 类型”的既有组织方式，失败选项文件同步使用类型名称命名。

## 基线问题

工作区开始时已有未跟踪的 `tests/Bing.Offices.EntityProviderContracts/Bing.Offices.EntityProviderContracts.xml`，本任务未修改。net6.0 依赖警告为现有工具链提示。

## 已知问题

无本轮新增问题。

## 风险与回归关注点

变更影响 Abstractions 公开源码位置及 NPOI 内部命名空间，风险评估为中等；程序集公开类型身份与运行路径未变。后续若有外部工具硬编码源码文件路径，需按新目录定位。

## Reviewer 注意事项

请以移动后的类型主体和原文件比较，避免把未暂存的删除/新增误判为类型删除。重点核对公开命名空间、动态列辅助类型的引用以及 NPOI 内部调用。

## Git 状态

未自动 `git add`、commit、push、创建 PR 或发布包。保留上述任务开始前的未跟踪 XML 文件。
