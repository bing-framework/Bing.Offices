# Public API Before/After Diff

基线为本轮修复前的 `HEAD` 源码，after 为当前工作树 Release 编译程序集。该文件只记录本轮 MUST_FIX 相关 API，不替代 API analyzer。

## Breaking

| Before | After | 迁移 |
| --- | --- | --- |
| `Bing.Offices.Settings.ExcelSetting.Default` 可替换静态实例 | 已移除；使用 `ExcelWorkbookMetadataOptions` | 在每个 Workbook 请求调用 `Metadata(...)` |

## Added

| API | 用途 |
| --- | --- |
| `ExcelWorkbookMetadataOptions` | 请求级六字段 metadata 快照 |
| `ExcelWorkbookExportBuilder.Metadata(...)` | 显式设置导出 metadata，并触发模板 override |
| `ExcelWorkbookExportRequest.Metadata` | 读取构建后的 metadata 快照 |
| `ExcelWorkbookExportRequest.MetadataSpecified` | 区分模板 preserve 与显式 override |
| `ExcelImportFailureOptions.TemporaryDirectory` | 请求级失败工作簿临时目录 |
| `CsvImportErrorCode` | CSV 错误分类 |
| `CsvImportError.Code` | 暴露 CSV 错误分类 |
| `CsvImportOptions<T>.MaxInputBytes` | CSV 输入字节上限 |
| `CsvImportOptions<T>.MaxRows` | CSV 数据行上限 |
| `CsvImportOptions<T>.MaxErrors` | CSV 错误数上限 |
| `CsvImportOptions<T>.MaxFieldLength` | CSV 单字段字符上限 |
| `CsvImportOptions<T>.MaxColumns` | CSV 单记录列数上限 |
| `CsvImportResult<T>.IsTruncated` / `MaxErrors` | 表达资源限制截断状态 |

## Machine Evidence

- Abstractions member snapshot before：`D1B4C608ECE5FC799F8C7704868A98241C6BB53BC4034FE455EB2AA937759186`
- Abstractions member snapshot after：`225DC5822857B4D660FC7944CB885BC46249D72C344E90B6654EDCC5DC1F15D9`
- Core member snapshot before：`40A788EE5B49AF9599942AB68DA946D924FF6062257F1831B5AEBCAA26D760BE`
- Core member snapshot after：`40A788EE5B49AF9599942AB68DA946D924FF6062257F1831B5AEBCAA26D760BE`
- NPOI member snapshot：`A0DBE9808D82547601429D8958C7ED283467031A3763EB9037B19D03F19D80BD`（无本轮 public member 变化）
- Public top-level type gate：新增 `Bing.Offices.Abstractions:Bing.Offices.Csv.CsvImportErrorCode`，已纳入批准清单。
- API 测试：`PublicApi_ReleaseAssemblies_ShouldMatchApprovedBaseline`、`PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`。
- 当前版本：`2.0.0`。
