# Entity 单据重复页眉与打印标题

Task ID：`BING-OFFICES-ENTITY-REPEATED-HEADERS-20261004-001`

## 目标

在 NPOI XLS/XLSX 和 ClosedXML XLSX 的 Entity Layout 中提供二选一的重复打印标题行与实际插入页眉行；实际页眉支持模板连续行及 Fluent 覆盖，导入能跳过并校验页眉，保留原始坐标。

## P1 公共契约

- `ExcelEntityLayoutBuilder<TEntity>.PrintTitleRows(sheetName, rows)` 只设置打印元数据。
- `ExcelEntityListRegionBuilder<TItem>.PageHeader(sourceRange, markerAddress, markerText, configure)` 要求 PageBreak；新增页眉上下文与构建器，支持相对单元格、值工厂、合并。实际页眉与同 Sheet 打印标题互斥。
- 首 页页眉区域连续且紧邻明细，包含已启用的列表表头；同 Sheet 至多一个实际重复页眉区域。公共布局描述提供不可变 Provider SPI，旧 API 不改签名。

## P2 Provider 与导入

- 共享行规划在分页符后、下页明细前插入页眉段；分页容量仅统计明细。两个 Provider 复制单元格内容、样式、行高和完整合并，随后按页执行 Fluent 覆盖。
- 模板与区域边界、冲突、marker 和名称在写入目标流前预检；导入扫描并跳过页眉，缺失、重复和错位报结构化错误，保留原始行列坐标。
- 打印模式复用 Provider 的现有重复标题行实现，保持无关模板设置。

## P3 验证与交付

- NPOI XLS/XLSX、ClosedXML XLSX 覆盖打印模式、模板/Fluent、分组与分页、命名锚点、空/单/多页、真实导入往返、失败及资源行为。
- 先职责测试，再受影响 Provider 双 TFM、NuGet 消费者、API 差异及 Release 常规门禁。
- 更新中文 XML 注释、文档、API 成员审批与快照、成员测试映射和 execution。无公共成员删除、版本号变化、自动提交、推送或发布。
