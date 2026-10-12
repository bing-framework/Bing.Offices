# Excel 能力补齐：ERP 字典模板与导入校验

Task ID: `BING-OFFICES-ERP-LOOKUP-TEMPLATES-20260928-001`

## 目标

完成 Workbook 与 Entity 模板的自动文本字典下拉，以及 NPOI/ClosedXML 名称列表导入校验。保持现有构造、接口和 Fluent 调用兼容；编码转换沿用 ValueMap/Converter。

## P1：公共契约

- Workbook 与 Entity 构建器新增 `LookupList(string name, IReadOnlyList<string> values)`；`ExcelDataValidationDefinition` 新增 `ListSourceName`，并使名称来源与显式 `Values` 互斥。
- Entity 构建器新增 `DataValidation(string sheetName, ExcelDataValidationDefinition validation)`；请求/布局通过只读、隐藏于 IntelliSense 的 SPI 暴露字典及 Entity 校验描述。
- `ExcelEntityImportOptions` 新增四参数构造重载，接收资源限制、失败策略、验证模式及不支持规则策略；旧重载继续只执行配置规则。
- 输入值立即快照，去首尾空白、拒绝空项、Ordinal 去重并保留首次顺序；名称大小写不敏感验证。

## P2：Provider

- 每个字典追加一个隐藏的 `__BingLookup_001` 顺序辅助 Sheet、一列纯文本和工作簿级名称；名称冲突在写入前拒绝，Sheet 名冲突顺延，不覆盖模板。
- NPOI XLS/XLSX 与 ClosedXML XLSX 写入名称来源下拉。NPOI 的显式列表 XML 引号修正仅修改显式列表。
- NPOI 对齐 ClosedXML 的名称列表导入规则：局部优先，拒绝失效局部、名称链、动态公式、二维/联合/外部引用及字典公式单元格；单次导入缓存允许值。
- Entity 固定单元格和实际明细行可选执行原生规则，顺序为原生、转换、配置；NPOI 临时区域维护原始坐标。失败值不进入结果，沿用现有错误与 Unique 回滚。
- 新增 `NamedListValidation` 细粒度能力声明；其他 Provider 对新模板请求预检拒绝。

## P3：验证与交付

- 公共模型与 Provider 职责级测试，真实采购/商品/盘点模板，完整结构与导入错误，NPOI 双格式及 ClosedXML XLSX。
- 同异步、取消、提交失败、旧目标文件、流所有权、临时文件和定向缓存资源证据；受影响项目双 TFM、NuGet 消费者、双 TFM API 差异及 `Category!=Large` 常规测试。
- 更新能力矩阵、可运行示例、迁移记录、API 审批、生产符号测试映射及 execution.md；检查 UTF-8、BOM、LF、`git diff --check`。
- 不修改版本号，不自动 commit、push 或发布。后续独立设计填写保护、级联字典、分批导入扩展及分组分页/透视报表。
