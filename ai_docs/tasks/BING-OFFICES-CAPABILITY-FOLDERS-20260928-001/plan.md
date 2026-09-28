# 按能力整理实体、导出和导入源码目录

Task ID: `BING-OFFICES-CAPABILITY-FOLDERS-20260928-001`

## 目标与兼容性

整理 Abstractions 中平铺的 `Entities`、`Exports` 和 `Imports` 目录，让相关类型相邻。所有公开类型保留现有命名空间、程序集、名称和签名；仅允许内部类型调整命名空间。此次只改文件组织和必要引用，不改业务逻辑。

## 实施安排

1. 实体布局：将布局与单元格、列表区域、计算列、动态列、Footer、分组小计、实体导入和实体导出分别归入子目录。内部地址解析、表达式及布局校验类型迁入对应能力命名空间；公开类型仍为 `Bing.Offices.Entities`。
2. 导出契约：按 Workbook、Sheets、DynamicColumns、Reports、Charts、Content、Streaming 归类。内部动态列辅助类型随 DynamicColumns 移动；公开类型仍为 `Bing.Offices.Exports`。
3. 导入契约：按 Workbook、Sheets、Batch、Failures、Validation、Resources 归类；通用入口留在 Imports。公开类型仍为 `Bing.Offices.Imports`。
4. Provider 对齐：将三个 NPOI 实体布局内部类型的命名空间统一为 `Bing.Offices.Npoi.Entities`，与已有 ClosedXML Provider 命名方式一致，并更新内部引用。不移动其他 Provider、Core 或 CSV 目录。
5. 使用文件移动和最小引用修改，保留 XML 注释与源码编码。检查项目文件及现行文档中的显式源码路径；历史任务记录不改写。

## 验收

- 逐类型核对移动前后声明和成员内容；除计划中的内部命名空间及引用外，不出现代码变化。
- Release 构建、`Category!=Large` 常规测试及受影响的双 TFM Provider 测试通过。
- 双 TFM 公共 API snapshot 成员差异为零；若出现公开成员变化，修正迁移，不更新基线掩盖差异。
- 检查文件归类、内部命名空间、`git diff --check`，并对改动 C# 文件执行严格 UTF-8 BOM、LF 和末尾换行检查。

## 默认约定

先整理三大能力目录，公开命名空间兼容优先。目录层级用于提高可查找性，不据此重命名公开 API；不提交、推送或发布。
