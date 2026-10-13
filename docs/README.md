# Bing.Offices 文档索引

本目录按当前仓库实现整理。功能结论以生产代码、公开契约和对应测试为依据；历史任务记录只作为验证证据，不作为当前能力声明。

## 快速入口

| 主题 | 文档 | 适用场景 |
| --- | --- | --- |
| 项目模块 | [项目模块与实现边界](project-modules.md) | 快速定位程序集、核心逻辑、依赖和限制 |
| 源码结构 | [源码目录与命名空间](source-layout.md) | 查找契约、Core、Provider 和测试代码 |
| Excel | [Excel 高级导入导出](excel/README.md) | Workbook、Entity Layout、模板、校验和 Provider 使用 |
| CSV | [CSV 导入导出](csv/README.md) | 流式 CSV、映射、资源限制和公式注入防护 |
| Analyzer | [Entity Layout 编译期诊断](excel/entity-layout-analyzers.md) | BOE/BOM 诊断规则与安装方式 |
| Provider 扩展 | [Entity Layout Provider 合同](excel/entity-provider-contracts.md) | 开发和验证第三方 Provider |
| AI 工作流 | [Universal Agent Workflow](ai-workflow/README.md) | 仓库内 Agent 计划、执行、审查和修复流程 |

## 功能文档

### 映射、导入与导出

- [Mapping Profile](excel/mapping-profile.md)
- [JSON/XML Mapping](excel/mapping-json-xml.md)
- [导入校验](excel/import-validation.md)
- [动态列与 Entity Layout](excel/dynamic-columns.md)
- [日期与 DateTimeOffset](excel/dates.md)
- [异常与 Observer](excel/exceptions-and-observers.md)
- [异步 IO 合同](excel/async-io.md)
- [自定义文件提交服务](excel/file-committers.md)

### Workbook 与 Provider

- [Provider 能力与边界](excel/09-providers.md)
- [第三方 Excel 引擎选型](excel/10-third-party-components.md)
- [NPOI 用户扩展](excel/npoi-extensions.md)
- [工作表保护](excel/sheet-protection.md)
- [两级联动字典模板](excel/cascade-lookups.md)
- [Workbook 透视表](excel/pivot-tables.md)
- [NuGet 迁移](excel/nuget-migration.md)

## 本次盘点结论

| 模块 | 实现状态 | 文档状态 |
| --- | --- | --- |
| Abstractions / Core | 已实现公共契约、映射计划、校验、CSV、文件提交和异常模型 | 已补充项目级模块说明；细节分散在 Excel、CSV 专题 |
| NPOI | 已实现 XLS/XLSX 主 Provider 和 Entity Layout | 已有专题与能力矩阵；已修正 XLSX 图表说明 |
| ClosedXML | 已实现富 XLSX Provider、Entity Layout 和透视能力 | 已有能力矩阵与专题 |
| MiniExcel | 已实现轻量 XLSX 导入导出 | 已有能力矩阵 |
| ExcelDataReader | 已实现 XLS/XLSX/XLSB 只读与批量导入 | 已有能力矩阵 |
| SpreadCheetah | 已实现前向流式 XLSX 导出 | 已有能力矩阵 |
| Aspose.Cells | 已实现商业渲染、转换、公式和加密适配 | 已有能力矩阵；部署许可证和目标环境保真度仍需单独确认 |
| Analyzer | 已实现 Entity Layout 与旧 API 迁移诊断 | 已有独立文档 |
| Provider 合同与测试工具 | 已实现可复用 Entity 合同、跨 Provider 合同和资源探针 | Entity 合同已有独立文档；仓库内部测试工具在模块总览中说明 |
| CSV | 已实现同步/异步导入导出、映射、校验和资源限制 | 本次新增独立专题 |

## 待实现 / 待确认

- **待确认**：NPOI 高级透视表在真实客户端中的文本 Count 等组合尚未完成最终验收，详见[透视表高级模式边界](excel/pivot-tables.md#高级模式与当前验收边界)。
- **待确认**：Aspose.Cells 需要有效商业许可证；字体替代、页面渲染和格式转换保真度应在实际部署环境验证。
- **待实现**：`templates/excel-provider-starter` 只覆盖基础 XLSX Provider 起点，尚未实现 Entity Layout；接入前不得声明对应能力。
- **待确认**：`docs/plans` 与 `ai_docs/tasks` 是计划和历史执行证据，不等同于当前产品能力；使用时应回到源码和本文索引复核。

## 维护约定

新增或调整公开能力时，应同步更新本索引、对应专题和 Provider 能力矩阵。示例优先引用公开接口，限制必须明确区分“不支持”“待实现”和“待确认”，不得根据计划文档推断功能已完成。
