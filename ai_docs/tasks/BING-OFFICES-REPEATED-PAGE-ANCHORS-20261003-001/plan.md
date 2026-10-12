# Entity 分页小计重复命名锚点

Task ID：`BING-OFFICES-REPEATED-PAGE-ANCHORS-20261003-001`

## 现状

`FooterNamed` 只为最终 Footer 写入单个工作簿名称。`PageSubtotal` 已能在 NPOI XLS/XLSX 和 ClosedXML XLSX 中逐页写 marker，但业务公式或外部程序无法按稳定名称定位各页小计。上阶段已完成分组与分页组合，行规划含有分页小计的准确行位置。

## P1 公共契约

- 新增 `ExcelEntityListRegionBuilder<TItem>.PageSubtotalNamed(string anchorPrefix, string markerText, Action<ExcelEntityListFooterBuilder<TItem>> configure = null)`，语义同 `PageSubtotal`，额外启用逐页名称。
- `ExcelEntityListRegion<TEntity>` 通过隐藏于 IntelliSense 的只读 `PageSubtotalAnchorPrefix` 向 Provider 传递描述；旧构造和旧 Fluent 调用保持兼容。`PageSubtotal` 再次调用时清除该前缀。
- 输出名称为 `{anchorPrefix}_{pageIndex:000}`，从 001 起，仅为存在下一页的完整页创建；超过 999 时按自然位数扩展。前缀和最终名称必须满足已有 Excel 名称校验与 255 字符上限。
- 布局阶段检查前缀与其他布局锚点的确定性冲突；运行时根据实际页数检查所有生成名称与模板名称，不允许大小写不敏感的歧义或跨 Sheet 占用。

## P2 Provider 行为

- NPOI 与 ClosedXML 在分页专用和分组分页组合路径都为每个分页小计 marker 写入名称；空明细与单页不生成名称。
- 模板中同名且属于目标 Sheet 的单格名称可重定位；其他 Sheet、公式、范围及同名多作用域冲突报 `Plan` 配置错误。
- 模板中属于同一前缀序列、但本次未生成的旧名称报错，避免数据缩短后留下过期页引用。无关模板名称和样式保持原样。
- 名称和行布局的验证在向目标流写入前完成；异常、取消和文件提交继续沿用现有合同。导入仍依据 marker，名称只服务导出定位。

## P3 验证与交付

- NPOI 的 XLS/XLSX、ClosedXML 的 XLSX 分别覆盖多页、组页同边界、空/单页、模板名称重定位、过期/冲突/无效名称、区域溢出和完整读回；NPOI 两格式均有成功和失败用例。
- 先跑职责测试，再跑受影响 Provider 双 TFM、公共 API 快照与包消费者，最后按影响范围执行常规门禁。
- 更新中文 XML 注释、Entity 文档、API 成员审批和双 TFM snapshot、迁移说明、生产符号到测试方法映射及 execution。只批准新增成员，不删除或改签名；不提交、推送、发布或修改版本号。

## 范围

本轮只定位分页小计 marker，不添加跨页重复单据头、打印标题、PivotTable 或公式引擎。重复页名称只代表有下一页的中间小计，不代表打印机实际分页。
