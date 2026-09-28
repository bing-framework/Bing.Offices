# Continuation 16：显式 Footer 公式单元格

## 目标与边界

当前 Footer 的 `Cell` 写入以 `=` 开头的字符串时，NPOI 输出文本，ClosedXML 输出公式。本阶段增加显式 `Formula` 方法，跨 Provider 稳定写入原生 Excel A1 公式，并保留 Footer marker、相对地址、样式、数字格式、导入跳过和目标流所有权语义。

公式文本由业务方提供，必须以 `=` 开头；框架不解析、改写或求值公式，也不提供公式 DSL。可变明细范围可由业务方使用 Excel 的 `ROW()`、`INDEX()` 等原生函数表达。已有 `PageSignatures` 签字区布局保持不变，本阶段不把公式与签字区打包为新抽象。

## 实施

1. 在 `ExcelEntityListFooterBuilder<TItem>` 增加 `Formula(string address, string formula, ExcelCellStyle style = null, string numberFormat = null)`。复用现有 Footer Cell 的地址、重复占位、合并和物理边界校验。
2. 增加隐藏于 IntelliSense 的公开公式值描述，用现有 Footer Cell SPI 的 `Evaluate` 返回它，不更改现有 SPI 接口签名。NPOI 与 ClosedXML 分别调用原生公式写入 API；显式数字格式优先于样式内格式。
3. NPOI HSSF/XSSF、ClosedXML XLSX 直接职责测试覆盖公式文本、数字格式、marker/导入边界、空明细和无效声明；同步与异步沿用共同布局执行器。
4. 更新文档、第三方包消费者、API 审批与双 TFM 快照、生产符号追溯和执行记录。

## Change Impact 与验证

- ChangedProjects：Abstractions、Npoi、ClosedXml、对应 Provider 测试、第三方消费者、文档与 API 快照。
- ChangedPublicContracts：新增一个 FooterBuilder 方法与一个 Provider SPI 描述类型；无成员删除或签名变更。
- RiskLevel：MEDIUM；重点是显式公式与普通字符串区分、跨 Provider 公式存储、样式和失败时输出保护。
- 顺序：职责测试 → 双 TFM Provider 定向/常规测试 → 包消费者 → API diff/审批 → Release 构建 → Docs/编码/diff 检查。`Category!=Large`；不运行无关性能矩阵。

## 后续

Footer 公式辅助构建器、旧版迁移分析器和显式 DOM/Streaming 策略继续作为独立 TODO。本阶段不修改 Utopa，不提交、推送或发布。
