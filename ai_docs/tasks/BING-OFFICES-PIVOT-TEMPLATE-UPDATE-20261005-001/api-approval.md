# 已有透视模板公共成员审批

- taskId: BING-OFFICES-PIVOT-TEMPLATE-UPDATE-20261005-001
- approvedBy: user（持续授权按 TODO 推进，采用本任务已批准 plan 的公共入口）
- approvedAt: 2026-10-05 (Asia/Shanghai)
- 范围：仅以下新增成员；不删除或改变既有类型、构造函数、签名、命名空间和默认行为。

## 精确成员

1. `Bing.Offices.Exports.ExcelWorkbookExportBuilder.PreserveTemplatePivotTables()`，返回当前 `ExcelWorkbookExportBuilder`。
2. `Bing.Offices.Exports.ExcelWorkbookExportBuilder.UpdatePivotTable(ExcelPivotTableDefinition definition)`，返回当前 `ExcelWorkbookExportBuilder`。
3. `Bing.Offices.Exports.ExcelWorkbookExportRequest.PreserveTemplatePivotTables`，公开只读 `bool`，`EditorBrowsable(Never)`。
4. `Bing.Offices.Exports.ExcelWorkbookExportRequest.PivotTableUpdates`，公开只读 `IReadOnlyList<ExcelPivotTableDefinition>`，`EditorBrowsable(Never)`。
5. `Bing.Offices.Providers.ExcelProviderFeatures.TemplatePivotTables = 1 << 26`，已在 NPOI XLSX 与 ClosedXML XLSX 的模板保留/更新主链及双 TFM 结构验收完成后添加并声明；该能力位不表示 Excel、WPS 或 LibreOffice 的刷新和显示结果已经验收。
6. `Bing.Offices.Exports.ExcelWorkbookExportBuilder.ResourceLimits(ExcelResourceLimits limits)`，设置透视模板预检的输入、ZIP 和 XML 资源限制；传入 null 时沿用默认值。
7. `Bing.Offices.Exports.ExcelWorkbookExportRequest.ResourceLimits`，公开只读 `ExcelResourceLimits`，`EditorBrowsable(Never)`，保存构建器的资源限制快照。

## 兼容与行为

旧构造、普通 Workbook 至少一个 Sheet 和普通新建透视规则保持兼容。保留/更新必须具备 XLSX 模板；仅显式模式允许模板请求无 AddSheet。更新隐含保留其它透视；名称与 TargetSheet 精确定位，不隐式新建。更新集合和定义深快照；构建失败不污染已生成请求。

P2 阶段 Provider 必须在枚举或输出前明确拒绝尚未实现的模式；不能以 API 存在或能力位替代 P3 输出链。模板输出主链完成后再移除对应拒绝和声明能力。

## 差异和候选验证

双 TFM 最终 API 差异预期包含上述七个成员。第六、七成员已进入当前候选，需重新捕获并比较双 TFM 快照；未经维护者成员级确认不更新正式基线。最终新包、源码/程序集身份、普通及第三方双 TFM消费者和门禁绑定本任务候选，不能复用 F3 冻结身份。

高级文本 Count、F-005～007、E-001/E-002 仍未完成。本文件是成员范围审批，不宣告实现或验证通过。
