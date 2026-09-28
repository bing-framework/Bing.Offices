# Entity Layout 连续明细求和公式

## 目标

在已有显式 `Footer.Formula` 之上，增加一个窄范围的公式辅助入口：为当前尾部上方连续的明细行生成原生 Excel 求和公式。空明细返回 `=0`，非空明细使用 `INDEX` 与 `ROW()` 定位实际行，避免硬编码首尾行号。NPOI XLS/XLSX 与 ClosedXML XLSX 使用现有公式值 SPI 写入，不增加 Provider 友元或公式解析器。

## 公共契约

在 `ExcelEntityListFooterBuilder<TItem>` 增加 `FormulaSumContiguousRowsAbove(string address, string sourceColumn, ExcelCellStyle style = null, string numberFormat = null)`。`address` 是尾部相对地址；`sourceColumn` 是绝对工作表列字母（A 到 XFD），必须是纯列字母。方法根据导出时该尾部收到的明细快照数量、最终 `GapRows` 和公式单元格的相对行生成引用。调用顺序不影响 `GapRows`；布局构建后保留快照隔离。

该方法可用于普通 Footer、分页小计和分组小计，其中传入的明细在尾部上方连续。最终 Footer 若同时配置分页小计或分组小计，布局构建阶段拒绝该方法，避免中间小计行被重复计入。原 `Formula`、`Cell` 和 Provider SPI 接口签名保持不变。其他聚合公式、非连续行公式及公式求值不在本阶段范围。

## 实施与验证

1. 构建器在 `Build` 时冻结 `GapRows` 并生成单元格定义；新增方法复用地址、重复占位、合并和物理边界校验。
2. 在列表区域构建阶段验证最终 Footer 与中间小计的非法组合，不限制分页小计和分组小计自身使用新方法。
3. 职责测试覆盖空/非空、表头、`GapRows` 前后调用顺序、相对下一行、样式与数字格式、HSSF/XSSF/XLSX 实际公式及读回、非法列和非法组合、目标流失败保护。
4. 更新第三方 NuGet 包消费者、中文 XML 注释、文档、双 TFM API approval/snapshot、公共类型分类（如需）和生产符号到测试方法映射。
5. 先跑职责级和受影响项目测试，再运行包消费者、API 检查、Release 构建与 `Category!=Large` 常规测试，最后执行严格 UTF-8/BOM/EOL 和 `git diff --check`。

## 范围

只修改 `Bing.Offices`；外部 Utopa 保持只读。不修改版本号，不自动 commit、push、发布或触发外部通知。跨非连续明细的公式范围生成、旧版迁移分析器和 DOM/Streaming 策略仍为后续独立任务。
