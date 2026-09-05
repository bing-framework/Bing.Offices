# Deprecated Removal

## 已完成迁移

| 旧符号 | 新符号 | 当前状态 |
| --- | --- | --- |
| `HeaderMatch` | `RequireExpectedHeaders` | 源码、测试、文档和 Benchmark 已迁移；旧成员删除 |
| `MaxColumnCount` | `MaxReadColumns` | 源码、测试、文档和 Benchmark 已迁移；旧成员删除 |
| `EnabledEmptyLine` | `ReportEmptyRows` | 源码、测试、文档和 Benchmark 已迁移；旧成员删除 |
| `IgnoreEmptyLineAfterData` | `StopAtFirstEmptyRow` | 源码、测试、文档和 Benchmark 已迁移；旧成员删除 |
| `MaxBytes` | `MaxSerializedBytes` | Failure Workbook 输出语义已迁移；旧 public 成员删除 |
| `AddNpoi` | `AddBingOfficesNpoi` | DI 入口、Docs、测试和 package consumer 已迁移；旧入口删除 |
| `ICellValueConverter` | `IExcelValueConverter` | legacy bridge、测试专用实现和接口已删除 |
| 六个旧 validation attributes | 对应 `Excel*` attributes | 生产识别、测试模型、Docs、API contract 和 consumer 已迁移并删除 |
| `CsvSeparatorCharacter` / `CsvQuoteCharacter` | 显式 delimiter/quote 参数 | 全局可变状态和依赖旧状态的旧重载已删除 |
| `_documentMappingConfiguration` | 无 | Import/Export Builder 死字段已删除 |

对旧名称的全仓扫描不包含任务历史报告中的迁移文字；生产源码中仍出现的 `maxBytes` 是内部流限制变量名，不是已删除的 public `MaxBytes` 成员。

## 保留或待批准项目

- `CsvHelper` DataTable 显式兼容类仍保留，当前仅删除全局状态和旧隐式重载；是否完全删除需要外部消费者与 breaking approval 证据。
- `OfficeException` 及 `OfficeHeaderException`、`OfficeEmptyLineException`、`OfficeDataConvertException` 尚未整体删除；仍有定义、API contract 或生产语义，需先完成错误分类迁移和外部异常捕获审计。
- `UniqueTracker`、剩余 Core `Execution detail` 类型和部分 Settings/兼容面仍需逐符号审批，不能以 `EditorBrowsable` 代替 internal 化。
- `ExcelMapping.For<T>()` 当前保留并按方向中立 Mapping configuration 入口治理，是否后续收敛另行审批。

## 扫描与结论

- 当前生产 `Obsolete` 命中：`0`。
- 当前生产 `TODO`：`0`。
- 生产程序集之间无 IVT；仅保留测试友元。
- 已批准清理子集：`VERIFIED`。
- 剩余兼容/DataTable/Execution detail 治理：`PARTIAL / BLOCKED`，不影响已完成迁移项的事实，但阻断最终 RC 发布。
