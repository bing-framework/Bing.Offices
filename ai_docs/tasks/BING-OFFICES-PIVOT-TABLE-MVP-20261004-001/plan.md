# Workbook 透视表最小能力

Task ID：`BING-OFFICES-PIVOT-TABLE-MVP-20261004-001`

## 目标

推进 DESIGN-006 中可在本地验证的公共 PivotTable 模型。Workbook 导出支持从已声明业务 Sheet 的明确有界区域建立一个行字段加一个求和字段的透视表，并将结果放到自动追加的独立 Sheet。ClosedXML XLSX 声明此能力；NPOI XLS/XLSX、MiniExcel 和 SpreadCheetah 在导出预检阶段明确拒绝。

## 实施

1. 在 Abstractions 增加 `ExcelPivotTableDefinition`、Workbook 构建入口、请求快照与细粒度能力位。源区域包含单行唯一非空表头和至少一行数据；目标 Sheet 不得与已有业务 Sheet 或其他透视目标重名。
2. ClosedXML XLSX 在业务 Sheet 完成写入后创建透视 Sheet 与 PivotTable，核对源区域、字段、目标地址及模板冲突。模板已有 PivotTable 仍沿用现有 Unsupported 边界；不承诺 Excel 公式计算或重新刷新外部数据。
3. NPOI XLS/XLSX、MiniExcel 和 SpreadCheetah 对本请求在写入前返回 Unsupported；NPOI 当前公开创建 API 不填充透视缓存记录，因此不声明伪实现。
4. 以真实 XLSX 检查透视缓存、字段和值；覆盖无效字段、重复名称、取消、目标文件保护和旧请求兼容。更新文档、双 TFM API 审批/快照、生产符号测试映射与执行记录。

## 边界

本阶段只支持 `Sum` 聚合，不支持列字段、筛选、多个值字段、分组、模板原有透视表保留或强制刷新缓存。公式重新计算及真实打印渲染另行做部署环境验收。不提交、推送、发布或修改版本号。
