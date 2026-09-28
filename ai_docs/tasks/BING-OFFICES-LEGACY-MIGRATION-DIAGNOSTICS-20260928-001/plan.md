# 旧版 Excel API 迁移诊断

Task ID: `BING-OFFICES-LEGACY-MIGRATION-DIAGNOSTICS-20260928-001`

## 依据与范围

只读业务样本 `D:\MyWorking\Utopa\Utopa.Erp.IE` 中，导出策略使用 `Bing.Offices.Exports.IExcelExportService` 和 `ExportOptions<T>`，商品模型使用 `Bing.Offices.Attributes.HasDynamicColumnAttribute`。前者应人工迁移到 `ExcelExport.Workbook` 请求及 `IExcelExporter`，后者的多个字典应在 Entity 列表区域显式配置 `DynamicColumnGroup`。

在可选的 `Bing.Offices.Analyzers` 包中新增两个 Warning：`BOM001` 提示旧导出服务/选项，`BOM002` 提示旧动态列标记。仅对已绑定到上述完整类型的语法报告，不凭相同简单名称推断；继续保留 2.x 的 `ColumnNameAttribute` 和 `DynamicColumnAttribute`。不提供 Code Fix，也不改变运行时程序集或版本。

业务自定义 `ImportLocationAttribute` 语义包含行列坐标、Sheet 选择和校验，不做通用诊断或自动转换；文档说明迁移到 `ExcelEntityCell` 与 `LayoutFromAttributes` 时需人工核对。

## 实施与验证

1. 新增独立的 `LegacyApiMigrationAnalyzer`，记录规则清单及中文迁移文档。
2. 双 TFM 职责测试覆盖旧类型引用、特性、2.x 保留类型、不同命名空间同名类型及无旧引用的代码。
3. 本地分析器 NuGet 包消费者增加条件编译的旧 API 样例，CI 验证真实打包分析器同时输出 `BOM001` 和 `BOM002`。
4. 更新主任务阶段索引与符号追溯；运行分析器测试、包消费者、差异和编码检查。

验收：两个诊断只命中完整旧类型；当前有效 2.x 属性不触发；双 TFM 测试和本地包消费者通过。外部 Utopa 仅只读，不修改、不复制其源码；不提交、推送、发布或改版本号。
