# 填写区域保护

Task ID: `BING-OFFICES-SHEET-PROTECTION-20260929-001`

## 目标

在 Workbook 与 Entity 模板中声明工作表保护和可填写区域。未配置保护的请求保持原行为；配置时，指定区域解锁，其余单元格沿用新工作簿或模板已有的锁定状态。工作表保护仅用于减少误编辑，不等同于文件加密或安全边界。

## 公共契约

- 新增 `ExcelSheetProtectionDefinition(IReadOnlyList<ExcelRangeDefinition> editableRanges, string password = null)`，构造时快照化区域，拒绝空集合引用与非法坐标；空区域集合表示整张工作表不可编辑。
- `ExcelSheetExportBuilder<T>.Protection(definition)` 和 `ExcelEntityLayoutBuilder<TEntity>.Protection(sheetName, definition)` 分别配置 Workbook/Entity 工作表。请求及布局公开只读的 Provider SPI 描述，保留全部旧构造及 Fluent 调用。
- 区域坐标为零基、闭区间，不随列表长度自动调整；可覆盖预留空白输入区。格式边界在 Provider 写入前预检。
- 增加 `SheetProtection` 细粒度能力位，只有实现了该能力的 Provider 声明支持。

## Provider

- NPOI XLS/HSSF、XLSX/XSSF 与 ClosedXML XLSX：内容、字典及原生校验完成后，按单元格/区域解锁指定输入区，再启用工作表保护。保留原样式中的字体、边框、格式及模板其他单元格的锁定设置。
- 空密码允许无密码工作表保护。模板已保护且未显式配置时不重设其保护；显式配置时以请求中的密码与填写区为准。
- MiniExcel 与 SpreadCheetah 对新保护请求在目标写入前抛结构化 Unsupported。ExcelDataReader 仅导入，不扩展导出能力。
- 同步与异步使用同一工作簿构建路径；取消、文件提交、目标保护和调用方流所有权保持现有合同。

## 验收

- 模型快照、非法坐标及格式边界；NPOI 双格式和 ClosedXML 的真实工作簿保护标志、输入区 Locked=false、非输入区 Locked=true、模板样式和原锁定状态。
- Workbook 与 Entity 各有直接测试；验证预留空白输入区、字典下拉共存、空密码与密码、未配置回归。
- 受影响项目双 TFM，API 成员新增差异、消费者示例、文档、生产符号测试映射、UTF-8/BOM/EOL 和 `git diff --check`。
- 不修改版本号，不自动提交、推送或发布。级联字典、分批导入和分组分页/透视报表仍在后续阶段。
