# Continuation 16 执行记录

## 结果

Entity Layout Footer 新增显式 `Formula` 方法。NPOI XLS/HSSF、XLSX/XSSF 与 ClosedXML XLSX 均将其写为原生公式单元格，支持现有 Footer 相对位置、数字格式、样式和 marker 导入边界。普通 `Cell` 方法和已有 Footer Cell SPI 签名未变。公式值通过隐藏于 IntelliSense 的公开描述类型传给 Provider，第三方 Provider 无需友元访问。

## Change Impact

- ChangedProjects：Abstractions、Npoi、ClosedXml、Provider 测试、API/Docs 测试、第三方包消费者、文档和任务记录。
- ChangedPublicContracts：新增 `ExcelEntityListFooterBuilder<TItem>.Formula`、`ExcelEntityFooterFormulaValue` 及其 `FormulaA1` 属性；无公共成员删除或签名变更。
- ChangedRuntimePaths：Footer 写入时识别显式公式值；普通单元格、导入执行、异步外围 IO 和文件提交路径保持原调用链。
- RiskLevel：MEDIUM；重点是双 Provider 公式存储、数字格式、空明细和失败时调用方流完整性。

## 验证

| 检查 | 结果 |
|---|---|
| NPOI Footer 公式职责测试 | net6.0/net8.0 各 7/7；HSSF/XSSF 真实读回、空/非空明细、格式、marker 边界、构建校验及无效公式失败保护 |
| ClosedXML Footer 公式职责测试 | net6.0/net8.0 各 3/3；XLSX 真实读回、空/非空明细、格式、marker 边界及构建校验 |
| Provider 常规测试 `Category!=Large` | NPOI net6.0/net8.0 各 696/696；ClosedXML 各 205/205 |
| Release 解决方案构建 | 0 错误；存量 net6.0 依赖支持与离线还原警告 |
| 第三方包消费者 | 从本地新打的 NuGet 包隔离还原；net6.0/net8.0 均运行成功并输出 `third-party-public-only-provider-ok` |
| API 快照 | 双 TFM 各只新增三个成员记录，`removed=[]`；审批文件和基线比较通过 |
| API 分类与文档代码块 | API 集成测试 net6.0/net8.0 各 32/32；Docs net8.0 10/10，18 个原文 C# 代码块执行 |

ClosedXML 允许保存语法不完整的 `=SUM(`，NPOI 会拒绝；因此本阶段只在布局构建时校验等号和非空公式体，不声称跨 Provider 解析或求值。NPOI 拒绝时验证目标流字节保持完整。先运行的全量常规测试暴露 API 分类与文档代码块计数遗漏；补齐后受影响项目均通过，其余项目首轮通过。外部 CI 尚未在本地运行。

## 生产符号追溯

`production-symbol-test-map.md` 的 Continuation 16 节记录构建器方法、Provider SPI 值对象、NPOI 和 ClosedXML 写入路径到测试方法的映射。

## TODO 分类

- Completed：显式 Footer 公式、双 Provider 实现、职责测试、第三方包消费、双 TFM API 审批与快照、文档和追溯。
- Open Actionable：无本阶段未完成项。
- Blocked Approval：无。
- Blocked External：远端 CI 结果尚不可得；本地验证不代替远端运行。
- Deferred：公式引用范围辅助构建器、旧版迁移分析器，以及有明确性能证据后的显式 DOM/Streaming 策略设计。
- Next Action：本阶段停止；后续按独立设计评估剩余 TODO。

IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
RELEASE_STATUS: NOT_EVALUATED
OPEN_ACTIONABLE: 0
