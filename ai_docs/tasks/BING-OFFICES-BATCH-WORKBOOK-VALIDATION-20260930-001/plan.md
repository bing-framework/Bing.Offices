# ExcelDataReader 分批 Workbook 原生校验

## 目标

为 ExcelDataReader 的固定列分批导入补齐 Workbook 原生 Data Validation，覆盖单 Sheet 批次和 Workbook 批次的同步、异步入口。保持完整导入、动态列、跨 Sheet 关系、失败工作簿和 XLS/XLSB 的现有边界不变。

## 实施范围

1. 在 ExcelDataReader Provider 内增加受限的 XLSX OOXML 校验读取器：解析工作表数据校验区域、工作簿/局部命名区域及一维单元格引用；按现有 Provider 契约支持列表、整数、十进制、日期、时间和文本长度规则。
2. 原生规则在配置转换前执行，错误使用 `WorkbookValidation`，保留行列坐标、空值策略、失败模式、取消和 UnsupportedFeaturePolicy 语义；命名列表按工作簿快照缓存。
3. 单 Sheet 分批请求使用默认 `Fail` 不支持策略；Workbook 分批请求使用请求中的 `UnsupportedFeaturePolicy`。规则不支持时按策略记录错误并决定是否接纳实体。
4. XLSX 之外的输入在启用原生校验时返回明确的 Provider Unsupported 预检异常；不修改完整 `Import` 的现有拒绝边界。
5. 增加 XLSX 真实工作簿职责测试：显式列表、名称列表、数值/文本长度、空值、同步/异步、Workbook 多 Sheet、Unsupported 规则和错误坐标；保留已有动态列及资源边界测试。

## 兼容性与风险

- 不新增公共 API，不修改现有请求类型和构造函数。
- 读取器只使用 `System.IO.Compression` 和安全 `XmlReader`，不引入 NPOI/ClosedXML 依赖。
- 只解析有界的一维区域和字面量边界；公式、外部引用、二维/联合区域和动态公式作为不支持特性处理。
- 原生校验快照与 ExcelDataReader 的顺序读取解耦，输入流位置在解析后恢复；取消和已有资源预检继续生效。

## 验证顺序

1. ExcelDataReader 原生校验职责级测试。
2. ExcelDataReader net6.0/net8.0 Release 构建与全量测试。
3. Docs 测试、`git diff --check`、严格 BOM/EOL/UTF-8 检查。
4. 记录最终生产符号到测试方法的映射；不提交、推送或发布。
