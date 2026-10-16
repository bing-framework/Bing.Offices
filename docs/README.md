# Bing.Offices 文档索引

本索引依据当前公开契约、实现及测试整理。完整的修订前状态、代码依据和未完成项见[功能与文档对照清单](documentation-inventory.md)；逐文件操作见[本轮文档变更](documentation-changes.md)。计划和历史记录不能代替当前能力证明。

## 入门与项目结构

| 文档 | 简介 |
| --- | --- |
| [快速开始](getting-started.md) | 包选择、服务注册及 Excel 导入导出往返示例 |
| [项目模块](project-modules.md) | 各程序集、Provider、合同工具的职责、依赖和限制 |
| [源码结构](source-layout.md) | 从契约、Core 和 Provider 目录定位实现 |
| [构建与验证](development/build-validation.md) | 构建、职责测试、文档测试、API、包消费者和资源探针 |
| [NuGet 迁移](excel/nuget-migration.md) | 包身份、真实替代接口和不兼容变更 |

## Excel 与 CSV

| 文档 | 简介 |
| --- | --- |
| [Excel 总览](excel/README.md) | Workbook、Entity、异步和 Provider 选择入口 |
| [工作簿、模板与报表](excel/workbook-reports.md) | 多表、元数据、报表内容、图片、图表与引擎差异 |
| [实体布局](excel/entity-layout.md) | 固定字段、多列表区域、命名锚点及模板基础用法 |
| [动态列与分页布局](excel/dynamic-columns.md) | 动态列、计算列、分组、小计、页眉和 Footer |
| [批量导入与失败工作簿](excel/batch-import.md) | 单表、多表、实体批次以及 Forward 和摘要模式边界 |
| [映射 Profile](excel/mapping-profile.md) | 导入导出方向、配置合并、命名转换器与计划缓存 |
| [JSON/XML 映射](excel/mapping-json-xml.md) | v2 文档、模型别名、加载入口和解析限制 |
| [导入校验](excel/import-validation.md) | 配置规则、原生校验、错误、资源限制与失败策略 |
| [日期与 DateTimeOffset](excel/dates.md) | 精确格式、1900/1904 日期系统和显式偏移策略 |
| [异常与观察器](excel/exceptions-and-observers.md) | 异常分类、注册及取消传播 |
| [异步 IO](excel/async-io.md) | 真实异步流、同步引擎边界与 DOM 资源准入 |
| [文件提交器](excel/file-committers.md) | 同目录暂存、原子提交、取消和自定义服务 |
| [工作表保护](excel/sheet-protection.md) | 保护定义、解锁区域与文件加密的区别 |
| [两级联动字典](excel/cascade-lookups.md) | 隐藏字典表、名称映射及下拉模板 |
| [透视表](excel/pivot-tables.md) | 基础/高级聚合、缓存、模板更新与客户端验收限制 |
| [公式、渲染与格式转换](excel/formula-rendering-conversion.md) | 公式读取、商业计算、PDF/图片、格式与加密转换 |
| [CSV 导入导出](csv/README.md) | 同步/异步流、映射、错误收集与公式注入防护 |

## Provider 与扩展开发

| 文档 | 简介 |
| --- | --- |
| [Provider 能力矩阵](excel/09-providers.md) | 六个引擎的实际接口、格式支持、注册与拒绝行为 |
| [第三方组件与版本](excel/10-third-party-components.md) | 当前依赖版本、用途及商业部署条件 |
| [NPOI 用户扩展](excel/npoi-extensions.md) | 单元格、工作表与服务注册扩展 |
| [布局与迁移分析器](excel/entity-layout-analyzers.md) | BOE/BOM 规则、安装、覆盖与限制 |
| [Entity Provider 合同](excel/entity-provider-contracts.md) | 第三方 Provider 的 SPI、合同包和接入流程 |
| [Provider 起始模板](../templates/excel-provider-starter/README.md) | 基础 XLSX Provider 模板，不含 Entity 实现 |

## 协作、夹具与历史资料

| 文档 | 简介 |
| --- | --- |
| [Agent 工作流索引](ai-workflow/README.md) | 计划、执行、审查、修复、目标生成及运行环境适配 |
| [Agent 规则片段](../templates/AGENTS.workflow.snippet.md) | 可复用的工作流规则示例 |
| [Golden 夹具](../tests/Bing.Offices.Testing/Resources/Golden/README.md) | 标准 Open XML 输入的离线生成与哈希维护 |
| [透视客户端探针](../build/PivotClientProbe/README.md) | Windows Excel 首开、刷新、重开及专用进程清理 |
| [中文注释历史计划](plans/chinese-comments-plan.md) | 历史治理计划，执行状态需回到源码核实 |

尚未实现或无法由代码确认的内容统一标为“待实现/待确认”，并注明原因。新增公开能力时同步更新专题、Provider 矩阵和本索引；原始任务、日志与客户端记录保留其时间语义。
