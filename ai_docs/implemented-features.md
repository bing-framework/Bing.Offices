# 已实现功能分类

更新日期：2026-10-20。按业务能力重新归类；实现依据来自当前 `src` 与现行 `docs`。本表说明库中已有的能力，不等同于所有 Provider、客户端或发布门禁均已通过。待实现和待确认事项集中在 [TODO](TODO.md)。

## 1. 基础接口与 Provider

| 已实现能力 | 实现位置 | 使用入口与边界 |
| --- | --- | --- |
| Excel/CSV 公共接口、Workbook/Entity 请求、能力声明与依赖注入 | `src/Bing.Offices.Abstractions`、`src/Bing.Offices.Core`、各 Provider 注册扩展 | [快速开始](../docs/getting-started.md)、[模块](../docs/project-modules.md)；不按接口存在推断引擎全部支持 |
| NPOI XLS/XLSX、ClosedXML XLSX、MiniExcel 轻量 XLSX | `src/Bing.Offices.Npoi`、`src/Bing.Offices.ClosedXml`、`src/Bing.Offices.MiniExcel` | [Provider 矩阵](../docs/excel/09-providers.md)；MiniExcel 不提供 Entity Layout |
| ExcelDataReader 只读 XLS/XLSX/XLSB、SpreadCheetah 前向 XLSX 导出 | `src/Bing.Offices.ExcelDataReader`、`src/Bing.Offices.SpreadCheetah` | 不把只读或流式 Provider 当成完整工作簿编辑器 |
| Aspose.Cells 商业计算、渲染及格式转换接口 | `src/Bing.Offices.AsposeCells` | 许可证、字体与目标环境另行验收，见第 8 类 |

## 2. 工作簿、模板与报表内容

| 已实现能力 | 实现位置 | 使用入口与边界 |
| --- | --- | --- |
| 单表/多表导入导出、模板和工作簿元数据 | Abstractions 的 `Exports/Workbook`、`Imports/Workbook` 与各 Provider | [工作簿与报表](../docs/excel/workbook-reports.md) |
| 表格、筛选、冻结、条件格式、名称、打印布局 | Abstractions 的 `Exports/Reports` 与 Provider 报表写入器 | 具体支持按 Provider 区分；打印布局写入不等于真实打印验收 |
| 图片、批注、原生数据校验、名称列表、保护与两级联动字典 | Abstractions 的 `Exports/Content`、`src/provider-shared/ExcelSheetContent.cs` 及 Provider 写入器 | [保护](../docs/excel/sheet-protection.md)、[级联字典](../docs/excel/cascade-lookups.md)；不宣称任意层级联动 |
| XLSX 柱状图、折线图和饼图 | Abstractions 的 `Exports/Charts`、NPOI 的 `Exports` | 仅 NPOI XLSX，不能外推到 XLS 或其他引擎 |

## 3. 实体布局、动态列与分页

| 已实现能力 | 实现位置 | 使用入口与边界 |
| --- | --- | --- |
| 固定字段、命名锚点、多列表明细、区域关系与模板布局 | Abstractions 的 `Entities`，NPOI/ClosedXML 的实体执行器 | [实体布局](../docs/excel/entity-layout.md) |
| 动态列、计算列、明细/Footer 求和、分组与分页小计 | `src/provider-shared/ExcelEntityRowPlan.cs` 与实体执行器 | [动态列与分页](../docs/excel/dynamic-columns.md) |
| 重复页眉、逐页名称与独立页锚点、导入边界检查 | `src/provider-shared/ExcelEntityPage*.cs` 等共享实现 | 不从锚点支持推断绘图复制、公式重定位或真实打印已完成 |

## 4. 批量导入、校验与失败输出

| 已实现能力 | 实现位置 | 使用入口与边界 |
| --- | --- | --- |
| 单表、Workbook、Entity 分批导入；Forward/投影/摘要相关实现 | Abstractions 的 `Imports/Batch`、`Entities/Import`，`src/provider-shared/Excel*BatchProjection.cs`、`ExcelForwardXlsxReader.cs` | [批量导入](../docs/excel/batch-import.md)；流式读取不代表整个链路恒定内存，组合预算仍待确认 |
| 配置规则、类型转换、唯一值、跨表关系及错误收集 | Core 的 `Validations`、`Mappings` 与 Provider 验证管道 | [导入校验](../docs/excel/import-validation.md) |
| 错误行工作簿与原工作簿批注输出 | NPOI/ClosedXML 导入器，ExcelDataReader 的失败工作簿写入器及共享前向写入器 | ExcelDataReader 的 AnnotatedOriginal 仅限受限 XLSX；客户端与资源边界见 TODO |

## 5. 映射、配置与日期

| 已实现能力 | 实现位置 | 使用入口与边界 |
| --- | --- | --- |
| 方向 Profile、Attribute/请求配置合并、命名转换器与映射计划缓存 | Core 的 `Configurations`、`Mappings` | [Profile](../docs/excel/mapping-profile.md) |
| JSON/XML v2 加载、模型别名与映射配置 | Core 的 `Configurations` | [JSON/XML](../docs/excel/mapping-json-xml.md)；不把历史未收口的审查直接标为现存缺陷 |
| 精确日期格式、1900/1904 日期系统与 DateTimeOffset 策略 | Core 的 `Dates`、`Mappings` | [日期](../docs/excel/dates.md)；DateOnly 统一公共合同仍为待实现/待确认 |

## 6. CSV、异步 IO 与文件安全提交

| 已实现能力 | 实现位置 | 使用入口与边界 |
| --- | --- | --- |
| CSV 同步/异步导入导出、映射、错误策略与公式注入防护 | Core 的 `Csv` | [CSV](../docs/csv/README.md)；不是单独的 Excel Provider |
| 流所有权、真实异步 IO、DOM 资源准入与取消传播 | Core 的 `IO`、NPOI 的 `IO` 及各 Provider | [异步 IO](../docs/excel/async-io.md)；引擎 DOM 操作仍有同步边界 |
| 同目录暂存、原子提交、自定义提交器、异常与观察器 | Core 的 `IO`、Abstractions 的 `Exceptions` 及异常边界 | [文件提交器](../docs/excel/file-committers.md)、[异常与观察器](../docs/excel/exceptions-and-observers.md) |

## 7. 透视报表

| 已实现能力 | 实现位置 | 使用入口与边界 |
| --- | --- | --- |
| 基础/高级透视定义、来源预检、缓存及布局写入 | Abstractions 的 `Exports/Reports`，NPOI/ClosedXML 的透视写入器 | [透视表](../docs/excel/pivot-tables.md)；高级文本 Count 部分组合尚未验收通过，不能当作完整可用能力 |
| 显式保留已有透视、受限模板更新及业务 Sheet 桥接 | `src/provider-shared/ExcelPivotTemplate*.cs` | 文件结构实现与客户端首开/刷新/重开分开验收；XLS/HSSF 不支持透视 |
| 客户端生命周期与资源探针 | `build/PivotClientProbe` | [历史验证摘要](references/pivot-client-validation.md)、[资源矩阵](references/pivot-resource-matrix.md)；历史成功不推广到新组合 |

## 8. 公式、渲染与格式转换

| 已实现能力 | 实现位置 | 使用入口与边界 |
| --- | --- | --- |
| 公式文本与缓存读取、独立公式处理接口 | Abstractions 的 `Formula` 及 Provider 公式处理器 | [公式与商业扩展](../docs/excel/formula-rendering-conversion.md)；NPOI/ClosedXML 不提供完整重算 |
| Aspose.Cells 公式计算、PDF/页面图片、格式转换及宏/加密配置 | `src/Bing.Offices.AsposeCells/Bing/Offices` | 有实现和配置入口；有效商业许可证、跨平台字体及目标格式保真仍待环境确认 |

## 9. 扩展接入、分析器与开发工具

| 已实现能力 | 实现位置 | 使用入口与边界 |
| --- | --- | --- |
| BOE/BOM 布局冲突与旧 API 诊断 | `src/Bing.Offices.Analyzers` | [分析器](../docs/excel/entity-layout-analyzers.md)；仅分析静态可确定的内容 |
| 第三方 Provider 合同、基础 XLSX Starter 与 NuGet 消费者 | `tests/Bing.Offices.EntityProviderContracts`、`templates/excel-provider-starter`、Consumer 项目 | [合同接入](../docs/excel/entity-provider-contracts.md)；Starter 尚未实现 Entity Layout |
| API 快照、职责测试、资源探针、文档示例及 Agent 工作流 | `build`、`tests`、`.agents`、`.github` | [构建与验证](../docs/development/build-validation.md)、[协作工作流](../docs/ai-workflow/README.md)；工具存在不等于发布/生产容量门禁已通过 |

具体接口、用法、依赖和配置以各专题为准。本次只整理资料，不改动业务实现，也不将清理动作记为功能重新验收。
