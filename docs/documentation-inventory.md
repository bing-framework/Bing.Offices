# 文档与实现对照清单

盘点日期：2026-10-15。范围为根 README、`docs`、Provider 模板与构建工具说明，以及 `src`、`tests`、`build`、`.agents` 中对应的实现。下表保留**修订前**的文档状态，入口指向本轮整理后的文档。

“文档完整”表示该主题已有功能、用法和限制说明；不表示所有 Provider 或所有客户端均支持。尚未实现和不能由当前代码证明的事项统一使用“待实现/待确认”，并说明原因。

## 功能逐项对照

| 功能模块 | 代码依据（仓库相对路径） | 修订前分类 | 本轮处理与文档入口 |
| --- | --- | --- | --- |
| 公共契约、依赖注入、Provider 选择 | `src/Bing.Offices.Abstractions`、各 Provider 的注册扩展 | 已实现但文档缺失或过时：没有完整入门示例 | 新增[快速开始](getting-started.md)，修订[模块总览](project-modules.md) |
| Workbook 多表导入导出、模板和元数据 | `Bing/Offices/Exports/Workbook`、`Imports/Workbook`（均在 `src/Bing.Offices.Abstractions` 下） | 已实现但文档缺失或过时：分散在总览和测试中 | 新增[工作簿与报表](excel/workbook-reports.md) |
| 表格、筛选、冻结、条件格式、名称范围、打印布局 | `src/Bing.Offices.Abstractions/Bing/Offices/Exports/Reports`、各 Provider 的报表写入器 | 已实现但文档缺失或过时：缺少集中入口 | 补充报表专题与 Provider 边界 |
| Sheet 图片、批注、原生校验、字典列表 | `src/Bing.Offices.Abstractions/Bing/Offices/Exports/Content`、`src/provider-shared/ExcelSheetContent.cs` | 已实现但文档缺失或过时：功能散落 | 补充报表专题，链接校验、级联与保护文档 |
| XLSX 柱状图、折线图、饼图 | `src/Bing.Offices.Npoi/Bing/Offices/Exports`、Abstractions 的 `Exports/Charts` | 已实现但文档缺失或过时：只有能力摘要 | 报表专题明确仅 NPOI XLSX 支持 |
| Profile、JSON/XML v2、合并与缓存 | `src/Bing.Offices.Core/Bing/Offices/{Configurations,Mappings}` | 已实现但文档缺失或过时：源码定位、依赖说明不足 | 完善[Profile](excel/mapping-profile.md)与[映射文档](excel/mapping-json-xml.md) |
| 类型转换与日期、DateTimeOffset | `src/Bing.Offices.Core/Bing/Offices/Dates`、`Mappings` | 已实现但文档缺失或过时：缺少调用和定位入口 | 补充[日期合同](excel/dates.md)关联的映射配置入口 |
| 配置校验、唯一值、Workbook 原生校验 | `src/Bing.Offices.Core/Bing/Offices/Validations`、各 Provider 的验证管道 | 已实现且文档完整：已有阶段、错误和上传示例 | 保留行为说明，补充[实现定位](excel/import-validation.md) |
| 固定单元格、命名锚点、多列表区域、关系 | `src/Bing.Offices.Abstractions/Bing/Offices/Entities`、NPOI/ClosedXML 的 `Entities` | 已实现但文档缺失或过时：基础布局与高级动态列混排 | 新增[实体布局](excel/entity-layout.md) |
| 动态列、计算列、小计、页眉、Footer、逐页名称 | `src/provider-shared/ExcelEntityRowPlan.cs`、各 Provider 的实体执行器 | 已实现且文档完整：已有较细边界和示例 | 保留[动态列与分页](excel/dynamic-columns.md)，补充定位与配置 |
| 单表、Workbook、Entity 分批导入 | Abstractions 的 `Imports/Batch`、`Entities/Import`，NPOI/ClosedXML/ExcelDataReader 实现 | 已实现但文档缺失或过时：缺少独立用户说明 | 新增[批量导入与失败工作簿](excel/batch-import.md) |
| ExcelDataReader 失败工作簿 | `src/Bing.Offices.ExcelDataReader/Bing/Offices/ExcelDataReader/ExcelDataReaderFailureWorkbookWriter.cs`、`ExcelDataReaderAnnotatedOriginalWriter.cs` | 文档描述与代码不符：根 README 和模块总览仍写“不支持” | 修正为 ErrorRowsOnly 与受限 XLSX AnnotatedOriginal，详见[Provider 矩阵](excel/09-providers.md) |
| CSV 同步/异步、映射、错误、注入防护 | `src/Bing.Offices.Core/Bing/Offices/Csv` | 已实现且文档完整：已有独立主题 | 核对并完善[CSV](csv/README.md)实现路径与用法 |
| 文件提交、取消、流所有权、异步 IO | `src/Bing.Offices.Core/Bing/Offices/IO`、NPOI 的 `IO` | 已实现但文档缺失或过时：文件提交器仍将已完成的 Starter/合同包列为后续工作 | 修订[文件提交器](excel/file-committers.md)、[异步 IO](excel/async-io.md) |
| 异常与观察器 | `src/Bing.Offices.Abstractions/Bing/Offices/Exceptions`、Core 的异常边界 | 已实现但文档缺失或过时：缺少用法与实现路径 | 完善[异常与观察器](excel/exceptions-and-observers.md) |
| 工作表保护、两级联动 | NPOI/ClosedXML 的保护与级联写入器 | 已实现且文档完整：已有坐标与拒绝行为 | 保留示例，补充[保护](excel/sheet-protection.md)、[级联](excel/cascade-lookups.md)的实现定位 |
| 基础/高级透视、模板保留与更新 | NPOI/ClosedXML 的 `Exports/PivotTables`、`src/provider-shared/ExcelPivotTemplate*.cs` | 已实现且文档完整：已有结构合同和客户端限制 | 保留[透视表](excel/pivot-tables.md)，区分库侧实现与客户端验收 |
| 公式文本、缓存读取、重新计算 | `src/Bing.Offices.Abstractions/Bing/Offices/Formula`、各公式处理器 | 已实现但文档缺失或过时：主要依赖能力矩阵 | 新增[公式、渲染与格式转换](excel/formula-rendering-conversion.md) |
| Aspose.Cells PDF、页面图片、格式/宏/加密转换 | `src/Bing.Offices.AsposeCells/Bing/Offices` | 已实现但文档缺失或过时：缺少集中配置用法 | 新增商业扩展专题；有效许可证和部署字体仍需确认 |
| NPOI 用户扩展 | `src/Bing.Offices.Npoi/Bing/Offices/Extensions` | 已实现但文档缺失或过时：缺少调用示例与文件定位 | 完善[NPOI 扩展](excel/npoi-extensions.md) |
| Roslyn 分析器 BOE/BOM | `src/Bing.Offices.Analyzers` | 已实现且文档完整：已有规则、安装和静态分析边界 | 保留规则表，补充[核心实现](excel/entity-layout-analyzers.md) |
| 第三方 Provider 合同、Starter、包消费者 | `tests/Bing.Offices.EntityProviderContracts`、`templates/excel-provider-starter` | 已实现但文档缺失或过时：合同包依赖、TFM 与接入说明不足，部分后续建议已落地 | 对齐[合同接入](excel/entity-provider-contracts.md)与模板说明 |
| 源码组织、构建、测试、API 快照和资源探针 | `Bing.Offices.sln`、`build`、`tests`、`.github/workflows/ci.yml` | 已实现但文档缺失或过时：文件数过时、缺少工具总览 | 更新[源码布局](source-layout.md)，新增[构建与验证](development/build-validation.md) |
| Agent 计划/执行/审查/修复工作流 | `.agents/{skills,scripts}`、`.github/{agents,prompts}` | 已实现但文档缺失或过时：英文标题/角色说明混杂，Goal 循环说明需对齐执行规则 | 更新并中文化[工作流文档](ai-workflow/README.md) |

## 待实现/待确认

| 项目 | 分类与可确认边界 | 对应入口 |
| --- | --- | --- |
| Starter 的 Entity Layout | **待实现/待确认（待实现）**：模板只提供基础 XLSX 导出，不能声明 Entity 接口能力 | [Starter](../templates/excel-provider-starter/README.md) |
| DateOnly 统一公共合同 | **待实现/待确认（待实现）**：未发现统一公开支持，不能按 DateTime 行为推断 | [日期](excel/dates.md) |
| NPOI/ClosedXML 完整公式计算 | **待实现/待确认（待实现）**：当前实现读取公式/缓存，重算请求拒绝；Aspose 有独立计算接口 | [公式](excel/formula-rendering-conversion.md) |
| 高级透视文本 Count 组合、模板更新后的客户端呈现 | **待实现/待确认（待确认）**：结构实现存在，历史 Excel 客户端记录含未通过组合；未据此宣称客户端全面可用 | [透视表边界](excel/pivot-tables.md#高级模式与当前验收边界) |
| 商业许可证、字体替换和跨平台渲染保真 | **待实现/待确认（待确认）**：代码配置存在，实际许可证与目标部署环境不由仓库证明 | [商业扩展](excel/formula-rendering-conversion.md) |
| NuGet 线上版本、当前候选发布与生产容量 | **待实现/待确认（待确认）**：仓库包配置、历史门禁和资源探针不能证明外部发布或任意生产规模 | [迁移](excel/nuget-migration.md)、[构建与验证](development/build-validation.md) |

没有某项 Provider 能力时，文档直接注明“不支持”；这与“产品全局尚未实现”不同。例如 NPOI XLS 不支持透视，不表示 XLSX 透视尚未实现。

## 文档范围与语言

- `docs/excel`、`docs/csv` 为现行功能文档；`docs/ai-workflow` 为仓库协作工具文档；入口统一在[文档索引](README.md)。
- 根 README、`build/PivotClientProbe/README.md`、Starter README、Golden 夹具 README 也纳入核对，避免只有 docs 内部正确。
- `docs/plans/chinese-comments-plan.md`、`ai_docs/tasks` 是计划及历史记录。历史任务、原始日志、客户端结果保持其时间语义，不改写成当前通过证明。
- 面向阅读者的英文标题、说明和角色介绍改为中文；公开 API 标识符、命令、路径、配置键、机器状态和第三方专有名称保留原文。
- 存量 `docs/ai-workflow/*.md`、部分 Excel 文档使用 CRLF；`dates.md` 原有混合行尾在本轮完整修订该文档时统一为默认 LF，其余按文件保留已有行尾，不做全仓格式转换。

本轮逐文件操作和摘要见[文档交付清单](documentation-changes.md)。
