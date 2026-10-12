# 公共 API 成员审批与迁移

- taskId: BING-OFFICES-ENTITY-LAYOUT-BATCH-20260930-001
- approvedBy: user（要求继续按 TODO 顺序实现实体布局分批导入）
- approvedAt: 2026-09-30 (Asia/Shanghai)
- 范围：仅增加实体布局分批导入选项、批次结果模型、隐藏 Provider SPI、8 个同步/异步普通及模板入口和 `EntityLayoutBatchImport` 能力位；不删除公共成员，不修改既有实体导入签名。

双 TFM API 快照各新增 44 个成员，删除 0 个。新增内容已由 `artifacts/api-compare/api-diff.json` 逐成员核对，并写入正式 baseline。

兼容性说明：既有 `IExcelEntityImporter`、实体布局构造和完整导入入口保持不变；NPOI 与 ClosedXML 通过新增 SPI 提供批次入口，其他 Provider 不声明该能力。
