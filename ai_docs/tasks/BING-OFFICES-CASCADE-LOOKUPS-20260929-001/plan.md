# ERP 级联字典填写模板

Task ID: `BING-OFFICES-CASCADE-LOOKUPS-20260929-001`

## 目标与边界

在现有 `LookupList`、工作簿原生列表校验和隐藏字典 Sheet 之上提供两级联动：同一行上级单元格的文本决定下级单元格的下拉候选。支持 Workbook 列表模板与 Entity 单据模板。外部业务系统在每次导出时提供字典快照；服务端导入的权威校验仍由既有配置规则维护，不执行 Excel 的 `INDIRECT/VLOOKUP` 公式，也不将级联公式误判为已支持的静态名称列表。

## P1 公共契约

- 增加 `ExcelCascadeListDefinition` Provider SPI，按调用时输入快照化上级文本与各子列表。上级键 trim 后不能为空、大小写不敏感唯一；子项沿用 `LookupList` 的文本长度、去空、Ordinal 去重约束。按上级键 Ordinal 排序，保证物理布局稳定。
- Workbook 与 Entity 构建器增加 `CascadeList(string name, IReadOnlyDictionary<string, IReadOnlyList<string>> values)`。`name` 同时是生成的上级列表名称，不能与普通字典或请求名称范围重名。
- `ExcelDataValidationDefinition` 增加 `CascadeSourceName` 和 `ParentColumnIndex`。列表来源在显式值、静态名称、级联来源三者中必须且只能选择一项。级联区域仅支持单列、同一行上级列；不允许父列位于下级区域内。校验定义快照保留新属性。
- Workbook 请求及 Entity 布局公开隐藏于 IntelliSense 的级联描述；新增细粒度能力位 `CascadeListValidation`。所有既有构造与调用保持兼容。

## P2 Provider 实现

- 每个级联字典追加一个隐藏辅助 Sheet：A 列上级显示文本、B 列对应子列表名称、C 列按上级顺序纵向保存全部子项。为 A、A:B 和每个子段建立工作簿级名称。内部 Sheet 和名称使用顺序编号，遇到模板同名对象时跳过；用户名称冲突则在写入目标前报结构化配置错误。所有字典单元格写为文本。
- 下级列表规则写入 `INDIRECT(VLOOKUP($A2,映射名称,2,FALSE))` 形式；列绝对、行相对，公式起始行与校验区域首行一致。未匹配上级值时 Excel 不提供下级候选。先建隐藏字典与名称，再应用校验，保留 Sheet 顺序和模板其他内容。
- NPOI HSSF/XSSF 与 ClosedXML XLSX 使用相同布局、范围上限和错误分类。验证子项累计行数、上级数量及公式长度不超过格式边界。同步/异步外围 IO、取消、文件提交、流所有权保持现有合同。
- MiniExcel、SpreadCheetah 对级联请求在写入前明确返回 Unsupported；ExcelDataReader 只导入，不声明该导出能力。

## P3 业务测试与兼容性

- 商品类别到产品、区域到仓库两组级联；包含中文、逗号、引号、换行及以 `=` 开头的文本，验证完整隐藏 Sheet、名称、子范围与下拉公式，模板样式和普通字典共存。
- NPOI XLS/XLSX、ClosedXML XLSX 均覆盖成功和格式越界/名称冲突失败；同步、异步、预取消、目标流与文件保持、模板流所有权。
- 模型覆盖输入快照、大小写冲突、空项、重复项、无效父列、多来源互斥及缺失级联来源。既有静态/内联列表调用继续通过。
- 包消费者在 net6.0/net8.0 通过 NuGet 使用新公开配置。生成双 TFM API 差异，只允许预期新增成员，不删除或改签名；维护成员审批及生产符号到测试方法映射。

## P4 文档与收口

更新字典模板示例、Provider 矩阵和迁移说明，明确 Excel 下拉用于填写体验，服务端业务规则仍由应用配置；原生级联公式在 `WorkbookRules` 导入模式按既有 Unsupported 策略处理。完成 Release 构建、`Category!=Large` 常规测试、Docs 测试、严格 UTF-8/BOM/LF、`git diff --check` 与执行记录。不修改外部 Utopa、版本号，不自动提交、推送或发布。

后续顺序仍为分批导入扩展，再评估分组分页组合与透视报表；不在本轮加入多级递归联动、跨行上级引用或公式求值引擎。
