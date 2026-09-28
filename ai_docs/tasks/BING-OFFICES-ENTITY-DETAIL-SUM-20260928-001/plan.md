# 跨小计明细总计公式

## 目标

在 Entity Layout 最终 Footer 增加公式辅助方法，按实际写出的明细行段生成原生 Excel 求和公式。分页或分组小计、间隔行和签字行均不进入求和范围；空明细写入 `=0`。NPOI XLS/XLSX 与 ClosedXML XLSX 使用同一公开 Provider SPI 语义。

## 实施

1. `ExcelEntityListFooterBuilder<TItem>` 新增 `FormulaSumDetailRowsAbove(address, sourceColumn, style, numberFormat)`。公开且隐藏于 IntelliSense 的值对象携带已校验的来源列，并基于一基闭区间明细行段生成公式。公式以 `SUM` 的多个区间引用明细；空集合使用 `0`，超过工作簿公式长度的布局明确失败。
2. 仅最终 Footer 可以声明此辅助方法。NPOI 和 ClosedXML 在分页、分组及普通列表的现有写出循环中跟踪真实明细行段，再写入公式。布局顺序、原有 `Formula`/`FormulaSumContiguousRowsAbove` 与导入 marker 行为保持不变。
3. 分别验证 NPOI HSSF/XSSF、ClosedXML XLSX 的分组与分页、多行小计/GapRows、空明细、公式文本、导入读回、错误边界和目标流保护；包消费者双 TFM 从 NuGet 包调用新公开能力。同步中文 XML、文档、API 成员审批、双 TFM 快照和生产符号追溯。

## 验证

先职责级测试，再受影响 Provider 与合同测试、第三方包消费者，最终 API 对比、Release 构建、`Category!=Large` 常规测试、Docs 与严格 UTF-8/BOM/EOL 检查。不跑无关大型性能矩阵。不修改外部 Utopa，不提交、推送、发布或修改版本号。
