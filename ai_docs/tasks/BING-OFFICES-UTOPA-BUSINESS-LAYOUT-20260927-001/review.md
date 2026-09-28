<!-- AI_REVIEW_STATUS: PASS -->
AI_TASK_ID: BING-OFFICES-UTOPA-BUSINESS-LAYOUT-20260927-001
AI_REVIEWED_AT: 2026-09-27T05:32:00.000Z

# 独立代码审查 Round 2

## 审查结论

第一轮发现的 `FIX-001` 至 `FIX-004` 已在当前工作树完成修复。第二轮复核检查了计划、执行记录、实际调用链、异常和资源边界、测试证据及 Git 差异，没有发现新的可执行问题。

```text
IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
PERFORMANCE_EVIDENCE_STATUS: NOT_APPLICABLE
RESOURCE_EVIDENCE_STATUS: PASS
EXTERNAL_GATE_STATUS: NOT_APPLICABLE
RELEASE_STATUS: PASS
GOAL_STATUS: COMPLETED
OPEN_ACTIONABLE: 0
```

## 验收矩阵

| 计划阶段 | 结论 | 证据 |
| --- | --- | --- |
| P1 公共实体布局抽象 | `PASS` | 属性式固定单元格、多个动态字典组、Footer、Provider SPI 均在 Abstractions 中有公开描述；旧 Fluent/单字典路径保留。 |
| P2 Provider 实现 | `PASS` | NPOI 动态列拒绝分支已移除；NPOI HSSF/XSSF 与 ClosedXML 的 Footer、marker、动态组和边界逻辑均接入真实执行器。 |
| P3 真实业务测试 | `PASS` | 商品档案三字典、赠品单空明细、盘点单多 Sheet、marker 错误、动态字典创建、Footer 物理边界及模板 GapRows 内容均有方法级测试。 |
| P4 API、文档与交付 | `PASS` | 中文 XML、Provider 文档、迁移文档、API approval、双 TFM 快照、Consumer、execution 和生产符号追溯均存在；本轮未改公共签名。 |

## 重点修复复核

### FIX-001：ClosedXML Footer GapRows 导入边界

`ClosedXmlEntityLayoutExecutor.ReadRegion` 现在使用 `markerRow - 1 - GapRows` 计算明细终点，并拒绝位于合法 gap 起点之前的 marker。`EntityTemplate_FooterGapWithTemplateContent_ShouldNotImportGapAsDetail` 用真实模板在 gap 行保留内容，确认该行不被物化为明细。

### FIX-002：动态字典创建异常分类

ClosedXML 的 `CreateDynamicDictionary` 统一检查目标类型、默认构造和 Setter，并将失败包装为 `BingOfficesConfigurationException` 的 `Plan` 阶段；接口字典仍自动创建。NPOI 的显式动态组分发保持同一结构化错误契约。自定义创建/Setter 的取消异常会透传，不被误报为配置错误。

### FIX-003：Footer 绝对物理边界

NPOI 使用 `IWorkbook.SpreadsheetVersion` 区分 HSSF/XSSF 的最后行列索引，ClosedXML 按 XLSX 物理上限预检 Footer marker、单元格和合并区域。预检发生在 Footer DOM 写入和目标流写出之前，越界统一为 `Plan` 配置错误，输出流保持空。

### FIX-004：P3 场景和执行证据

新增测试方法已写入两个 Provider 测试项目，并登记到 `production-symbol-test-map.md`。`execution.md` 的 Round 1 记录、测试计数和后续取消语义说明与当前实现一致。

## 验证证据

| 检查 | 结果 |
| --- | --- |
| NPOI Entity Layout Provider 直接职责筛选，net6.0/net8.0 | `34/34`、`34/34` 通过 |
| ClosedXML Entity Layout 与 gap 模板直接职责筛选，net6.0/net8.0 | `15/15`、`15/15` 通过 |
| NPOI Provider `Category!=Large`，net6.0/net8.0 | `645/645`、`645/645` 通过 |
| ClosedXML Provider `Category!=Large`，net6.0/net8.0 | `168/168`、`168/168` 通过 |
| Provider Contract，net6.0/net8.0 | `276/276`、`276/276` 通过；取消边界后相关源码范围未变，复用该证据 |
| Release Solution Build | 0 错误；4 个既有 `Microsoft.Bcl.Memory` net6.0 警告 |
| UTF-8/BOM/LF 检查 | 修改的 C# 为 UTF-8 BOM + LF，任务文档为 UTF-8 + LF |
| `git diff --check` | 无新增空白错误；仅有 `ProfileFixtures.xml` 存量 CRLF 提示 |
| `InternalsVisibleTo` | 生产程序集只指向对应测试程序集，未发现 Provider 生产友元 |
| API / NuGet Consumer | 公共源码范围未受本轮私有修复影响，沿用前轮双 TFM API/Consumer 通过证据 |

## 追溯与范围

- 生产符号到测试方法映射：`ai_docs/tasks/BING-OFFICES-UTOPA-BUSINESS-LAYOUT-20260927-001/production-symbol-test-map.md`。
- 执行记录：`ai_docs/tasks/BING-OFFICES-UTOPA-BUSINESS-LAYOUT-20260927-001/execution.md`。
- 外部 `D:\MyWorking\Utopa\Utopa.Erp.IE` 未修改、未复制、未提交。
- 未执行 commit、push、PR、发布或版本号修改。

## 剩余分类

### Completed

计划 P1-P4 与第一轮 Review 的四项 `SHOULD_FIX` 均完成并有验证证据。

### Open Actionable

无。

### Blocked Approval

无。

### Blocked External

无；计划未要求外部 CI 或生产环境门禁。

### Not Applicable

无关的 500K/1M 性能矩阵、外部生产探针和外部 CI 不属于本轮变更验证范围。

### Accepted Limitations

MiniExcel、SpreadCheetah、ExcelDataReader、AsposeCells 不声明 Entity Layout 能力，继续按能力矩阵拒绝或不注册；公式 DSL、命名锚点、分页小计、契约测试包、分析器和 Source Generator 按计划延期。

### Deferred

计划表中 P2/P3 后续优化建议保持延期，不影响当前验收。
