<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-GROUP-PAGINATION-20261003-001
AI_EXECUTION_FINISHED_AT: 2026-10-03T14:59:04Z

# 实施执行报告

## 已完成

- 允许 Entity 列表区域组合连续分组小计、水平分页和可选分页小计，保留原有公开成员与签名。
- 两个 Provider 共用内部行规划和组合小计导入扫描，按明细数量分页；组页同边界依次写组小计、页小计和分页符。
- NPOI XLS/XLSX 与 ClosedXML XLSX 均支持组合导出、回读、原始坐标、空明细、无页小计、多行尾部和区域预检。
- 在组合中间小计中拒绝会跨入其他小计行的连续行求和公式；最终 Footer 仍可使用明细行范围公式。
- 更新 Entity Layout 文档示例、迁移说明和文档示例数量断言。生产符号到测试方法见 `symbol-test-map.md`。

## 验证

- `dotnet build Bing.Offices.sln -c Release --no-restore -m:1 --verbosity quiet`：通过，0 错误；4 个既有 Microsoft.Bcl.Memory/net6.0 支持警告。
- NPOI 职责测试：net6.0 与 net8.0 各 17/17；ClosedXML 职责测试：各 9/9。
- NPOI 常规测试：net6.0 与 net8.0 各 835/835；ClosedXML：各 271/271。Provider 合同：各 298/298。
- 全解决方案 `Category!=Large` 首次运行中，除 Docs 测试外全部通过。Docs 测试因新增一个 C# 示例使固定数量断言由 20 变成 21 而失败；修正断言后独立重跑 Docs 19/19 通过，新增示例已编译并执行。
- 公共 API 快照测试在 net6.0、net8.0 各 1/1 通过；本阶段未更新 API 基线。
- `git diff --check`、严格 UTF-8/BOM/EOL 和末尾换行检查通过；两个 Provider csproj 已按仓库规则统一为 CRLF。

## 范围

未修改版本号，未提交、推送或发布。重复页命名锚点、透视表和公式 DSL 留待独立设计。
