<!-- AI_REVIEW_STATUS: PASS -->
AI_TASK_ID: BING-OFFICES-PIVOT-TABLE-MVP-20261004-001
AI_REVIEWED_AT: 2026-10-05T02:22:52+08:00

# Round 1 最终独立复核

## 结论

本次对 `review.md`、`review-progress-20261005.md`、`review-progress-20261005-round2.md` 中 FIX-001～008 做独立复核。当前代码、直接测试、符号映射、TRX 与 API 差异产物相互一致，八项 SHOULD_FIX 均已解决；未发现本轮修复范围内仍未解决的 MUST_FIX 或 SHOULD_FIX，因此本轮状态为 PASS。

该结论只表示 PivotTable MVP 的既定范围及 Review Round 1 已收口，不表示长期 Excel 能力 Goal 全部完成。公式隐式计算行为已明确登记到 `BING-OFFICES-PIVOT-SOURCE-CONTRACT-20261005-001`，本报告不把它伪造为已解决，也不将该后续计划倒灌为当前 FIX。

## FIX-001～008 复核

| FIX | 结论 | 实现与直接证据 |
| --- | --- | --- |
| FIX-001 固定业务来源名单 | PASS | `ExcelWorkbookExportBuilder.Build` 在加入透视目标前复制 `sourceNames`，来源校验不再受后续目标名称污染；顺序隔离和多透视共源由 `PivotSource_ShouldRejectRegardlessOfDefinitionOrder`、`MultiplePivots_ShouldPreserveBusinessSourceAndAllDefinitions` 直接覆盖。 |
| FIX-002 工作表名称预检 | PASS | `ExcelSheetNameValidation.Validate` 集中处理长度、七种禁止字符、首尾单引号、NUL 与 U+0003；`ExcelPivotTableDefinition.Validate` 在业务行枚举前调用。非法目标和合法中文/大小写路径均有直接测试。 |
| FIX-003 透视展开边界 | PASS | `ClosedXmlReportWriter.Validate` 使用 `long` 校验两列宽度和静态边界；`ApplyPivotTables` 按实际去重分组数校验表头、分组行与总计空间。末行、末列、`int.MaxValue`、重复/空键和精确边界均有直接测试。 |
| FIX-004 职责及资源矩阵 | PASS | 同步/异步完整缓存、OOXML 关系、调用方流、真实文件成功、配置失败、写入中取消、序列化后取消、提交失败、旧目标保护、临时文件清理与单次异常观察均有职责测试。透视筛选产物双 TFM 各 47/47。 |
| FIX-005 XML 注释格式 | PASS | 受影响文件的 XML 文档块已恢复连续结构；最终 Release 构建记录为 0 错误，仅保留 4 条既有 `Microsoft.Bcl.Memory` net6.0 提示，没有新增 CS1570。 |
| FIX-006 辅助 Sheet 名称预留 | PASS | `ClosedXmlExcelExporter` 将全部透视目标传给普通与级联辅助 Sheet 分配器；二者按大小写不敏感规则同时避开已有和预留名称。普通/级联、多透视、大小写及模板已有辅助 Sheet 的组合测试均已接入。 |
| FIX-007 Entity 名称入口直测 | PASS | `EntitySheetNameValidationTest` 通过固定单元格、列表和合并三个公开入口验证 12 个非法名称及 2 个合法名称，共 6 个测试方法、42 个案例；合法路径核对完整关键布局描述。双 TFM 各 42/42，合同全套双 TFM 各 391/391。`symbol-test-map.md` 已映射共享校验、实体构建入口和六个测试方法。 |
| FIX-008 执行与证据收口 | PASS | `execution.md` 已标记 COMPLETED，包含实际完成时间、历史失败原因、最终构建、测试、包消费者、API、编码与符号追溯结果；历史 review 保持原样。 |

## 产物核验

- `artifacts/review-fix-tests`：2 份 TRX，net6.0、net8.0 各 47/47，合计 94/94，0 失败、0 跳过。
- `artifacts/entity-name-direct-tests`：2 份 TRX，双 TFM 各 42/42，合计 84/84，0 失败、0 跳过。
- `artifacts/contracts-final-tests`：2 份 TRX，双 TFM 各 391/391，合计 782/782，0 失败、0 跳过。
- `artifacts/review-fix-regular-tests`：27 份 TRX、14 个测试项目，合计 4,289/4,289，0 失败、0 跳过；没有把独立重跑拼接为全量结果。
- `artifacts/review-fix-api/api-member-diff.json` 与 `artifacts/review-fix-api-verified/api-diff.json`：net6.0、net8.0 均为空差异。该证据表示 Review 修复候选相对已批准基线没有新增或删除成员，不否认 MVP 初始批准的 14 个新增成员。
- 八个 2.0.0 生产包及隔离消费者产物存在；execution 登记 net6.0 与 net8.0 消费者均输出 `package-consumer-ok`。最终构建退出码为 0，0 错误、4 条既有提示。
- 本次复核执行 `git diff --check`，退出码为 0；仅报告既有 ProfileFixtures XML 的 CRLF 工作区提示。

## 范围边界

本轮没有重复运行全量测试，结论基于当前源码、直接测试实现和已保存的可解析 TRX/API 产物。高级透视、多 Provider 创建能力、模板透视保留、低内存导入、失败工作簿扩展、商业公式与渲染验收仍属于独立后续范围。公式来源值读取会触发计算的已知行为已由 SOURCE-CONTRACT 计划接管，当前 MVP 报告仅如实保留该限制。
