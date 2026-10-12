# API 边界与输入行为迁移

- taskId: `BING-OFFICES-PIVOT-SOURCE-CONTRACT-20261005-001`
- approvedBy: user（继续按 TODO 建议顺序实现）
- approvedAt: 2026-10-05（Asia/Shanghai）
- 范围：内部来源预检、结构化诊断、职责测试及公开契约说明。

## 成员边界

本阶段不新增、删除或修改任何公共成员签名。新增 `ClosedXmlPivotSourceValidator` 是 ClosedXML Provider 内部类型，按能力置于 `Exports/PivotTables`；复用现有带坐标的配置异常构造，不增加生产友元。

双 TFM 候选必须先对既有批准基线证明成员差异为零，才能刷新源码、程序集及包身份。原基线的成员审批元数据保留，不以覆盖成员快照掩盖差异。

## 输入行为迁移

纯值来源保持兼容。来源内公式及 Excel 错误在任何来源取值或缓存创建前拒绝，返回 `ConfigurationInvalid/Preflight`，包含透视名称和来源工作表的一基行列。模板公式无论缓存新鲜、过期或缺失均执行相同规则。

需要透视公式结果时，先通过应用或显式计算接口求值，再写为字面值。数字文本、日期及布尔求和值不隐式转为数值。金额双精度和舍入仍由应用维护。来源范围外的公式保持原保存契约。

完整示例与限制见 `docs/excel/pivot-tables.md`，迁移说明见 `docs/excel/nuget-migration.md`。该输入行为变更由来源合同 plan 明确规定，不新增公式计算或缓存读取选项。
