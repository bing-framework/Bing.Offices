# 公共 API 成员审批与迁移

- taskId: BING-OFFICES-ENTITY-INDEPENDENT-PAGE-ANCHORS-20261005-001
- approvedBy: user（持续要求按建议顺序实现 TODO；F-001 为独立逐页页眉与签字区锚点）
- approvedAt: 2026-10-05 (Asia/Shanghai)
- 范围：只新增本计划列出的成员；不删除或改变既有公开签名、默认值及行为。

## 批准新增成员

- ExcelEntityListRegionBuilder<TItem>.PageHeaderNamed：前缀、源范围、marker 地址/文本、可选配置委托及默认 Reject 的过期名称策略，精确参数见 plan.md。
- ExcelEntityListRegionBuilder<TItem>.PageFooter：marker 文本及可选现有 Footer 配置委托。
- ExcelEntityListRegionBuilder<TItem>.PageFooterNamed：前缀、marker 文本、可选配置委托及默认 Reject 的过期名称策略。
- ExcelEntityListRegion<TEntity>.PageHeaderAnchorPrefix。
- ExcelEntityListRegion<TEntity>.PageHeaderAnchorStaleNamePolicy。
- ExcelEntityListRegion<TEntity>.PageFooter。
- ExcelEntityListRegion<TEntity>.PageFooterAnchorPrefix。
- ExcelEntityListRegion<TEntity>.PageFooterAnchorStaleNamePolicy。
- ExcelEntityPageAnchorStaleNamePolicy 枚举类型及 Reject、RemoveObsolete 枚举值。

五个只读描述属性为 EditorBrowsable(Never) 的 Provider SPI；公开命名空间、程序集及旧成员保持兼容。枚举文件独立归入 Entities/Anchors。

## 行为与迁移

旧 PageSubtotalNamed 继续只生成中间页小计名称，旧 FooterNamed 继续定位最终 Footer；不自动升级为新系列。要定位每个实际页眉或每页签字区，显式使用新入口；空明细视为一页，末页也有独立 Footer。

新策略默认 Reject，模板缩短后旧系列名称仍会导致预检失败。只有显式选择 RemoveObsolete 才清理当前系列中合法、指向目标 Sheet 的过期单格名称，不清理歧义、非法、跨 Sheet 或无关名称，不改输入模板。

## API 验证状态

已保留更新前基线 artifacts/api-baseline-before.json，并从实际程序集生成双 TFM 候选。artifacts/api-member-diff.json 中 net6.0/net8.0 均仅有 Abstractions 新增 12 条、删除 0 条，其它七个生产程序集成员差异为零；新增参数、默认 Reject 策略和五项 EditorBrowsable(Never) 与上述批准范围一致。

12 条包括三个 Fluent 方法、五个 SPI 属性，以及枚举类型、Reject、RemoveObsolete 和编译器生成的 value__ 字段；value__ 是枚举固有表示，不是额外设计的公共扩展。旧构造、方法参数及默认值均无删除或签名变化。完整新增成员文本保留在差异 JSON，不以替换基线掩盖变化。

首次 api-compare-approved 已通过，随后才更新仓库快照；原始成员差异和更新前基线保留。包消费者与常规回归独立登记，不由 API 成员差异替代。

## 组合验收发现的内部修复

NPOI 实体区域原生校验读取动态列和计算列不存在的 Property 时会抛空引用。本阶段真实组合职责发现后，改为跳过不回写实体的计算列，并为动态列使用其 Key 作为错误字段键，与 ClosedXML 的既有校验顺序一致。没有新增或改变公共签名，也不改变原生校验失败后跳过整行的合同。

生产修复导致候选身份变化，须在最终构建、职责和回归通过后重新打包、捕获并比较；本次应相对已批准成员快照保持零成员差异。先保留修复前快照和候选，再验证新批准候选身份，比较通过后才更新身份元数据，不增加本计划之外的公开成员。
