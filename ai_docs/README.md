# ai_docs 索引

更新日期：2026-10-20。

| 入口 | 内容 |
| --- | --- |
| [已实现功能分类](implemented-features.md) | 按 9 类能力定位当前实现、使用文档与支持边界 |
| [统一 TODO](TODO.md) | 原 29 组未收口任务合并后的待实现/待确认事项、下一步和关闭条件 |
| [现行功能文档](../docs/README.md) | 接口用法、配置、依赖、示例和限制 |
| [透视客户端历史摘要](references/pivot-client-validation.md) | 基础成功与高级未通过的适用边界，不充当本轮验收 |
| [透视历史资源矩阵](references/pivot-resource-matrix.md) | 保留必要测量数值；原始产物已清理 |

## 资料整理结果

前一轮已清理 35 个已完成任务的过程资料。本轮把剩余 29 组未收口任务合并为一份 TODO，清理旧计划、执行记录、审查往返、旧 TODO 和生成产物；旧任务编号仅作为追溯标识保留在新清单中。已完成实现按功能归类，不再按历次任务重复堆放。

## 构建兼容审批附件

`tasks` 下仅保留以下 3 份被现有工具直接引用的审批文件，它们不是旧 TODO，也不表示相应任务全部通过。保留原路径与内容，避免改变审批身份、现有 API 基线和 CI 默认参数。

| 审批附件 | 当前引用 |
| --- | --- |
| [公共 Provider 边界](tasks/BING-OFFICES-PUBLIC-PROVIDER-BOUNDARY-20260927-001/api-approval.md) | `.github/workflows/ci.yml` |
| [报表及 Provider API](tasks/BING-OFFICES-ADVANCED-EXCEL-ROADMAP-20260926-001/api-approval.md) | `build/ApiSnapshot/Program.cs` 默认审批路径及前一审批的引用 |
| [NPOI 透视 API](tasks/BING-OFFICES-NPOI-PIVOT-CACHE-20261005-001/api-approval.md) | `build/api-snapshot-baseline.json` 的审批身份路径 |

## 后续维护

待办仅维护 `TODO.md`；历史结论未经当前复核的条目标为“待确认”，不得直接当作当前缺陷或已通过证据。本轮只整理资料，不执行这些待办。

需要实施某一待办时，再使用现有工作流在 `tasks/<taskId>/` 下建立对应计划。完成后将稳定用法更新到 `docs`、同步功能分类并关闭 TODO；任务过程资料及时清理。对外部审批、客户端和资源门禁分别记录，不以源码存在代替验收通过。
