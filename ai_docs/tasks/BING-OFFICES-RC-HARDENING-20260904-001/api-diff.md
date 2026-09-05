# API Diff

## 结论

`BLOCKED / 待批准`。当前 Release candidate 与上一正式 API baseline 存在本任务明确引入的异常合同、日期配置、资源限制和 NPOI Provider User API 变化；正式 baseline 未修改。

## 快照与哈希

- 正式 baseline：`ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/artifacts/api-snapshot-formal-baseline-20260903.json`
- 本任务 candidate：`ai_docs/tasks/BING-OFFICES-RC-HARDENING-20260904-001/artifacts/api-snapshot-candidate-20260904/`
- candidate 快照覆盖：`netcoreapp3.1`、`net6.0`、`net8.0`。
- candidate hashes：
  - `Bing.Offices.Abstractions`: `F66D947DF7B3DA0D6C7C1AE7790FC6BE5A77BDFE8319ACCDB79B3AD3D4D44BC3`
  - `Bing.Offices.Core`: `AE897589BA23D2DAF762DFA10A468CCA32A783BF45F5D0F772C36E7A73FEE9DC`
  - `Bing.Offices.Npoi`: `1C5E5DCEA0BDDA945ABCEED01B4DFFE24EE98593032D306C86FF4E61FECC7427`
- formal hashes：
  - `Bing.Offices.Abstractions`: `7F9A2AA819E94B3838097DF2FF374A934CF7F35F3D2E91F3D1DB790F22972943`
  - `Bing.Offices.Core`: `B3661970BBE5AECC06DAD57B1E3F960FA77E70C4D2E66B2DA4910F7823AA2BB6`
  - `Bing.Offices.Npoi`: `DA163263804A964D8AC2A13D78D6B3858256171CE7729841690FDB56F602CEEE`

candidate 生成器的 `memberCount` 为 Abstractions `771`、Core `249`、NPOI `73`；本报告不把不同快照版本的 line 表示数量与逻辑集合数量混用，成员级原始数据以 candidate JSON 和 formal baseline 为准。三个 candidate TFM 的哈希和公共成员形态一致。

## 语义变化

### Abstractions

- 新增统一异常合同：`BingOfficesException`、Configuration/Import/Export/ResourceLimit/FileCommit/Unsupported 子类及 `BingOfficesErrorCode`、`BingOfficesOperation`、`BingOfficesStage`。
- 新增 `ExcelCellValue.IsDate1904`，用于 1900/1904 serial 语义。
- `ExcelResourceLimits` 新增 ZIP entry 数、单 entry/总解压大小、压缩比、`sharedStrings.xml`、`styles.xml`、单 worksheet 和 worksheet 总量限制。
- `ExcelImportOptions`/Sheet request 的命名迁移：`HeaderMatch`、`MaxColumnCount`、`EnabledEmptyLine`、`IgnoreEmptyLineAfterData` 分别迁移为 `RequireExpectedHeaders`、`MaxReadColumns`、`ReportEmptyRows`、`StopAtFirstEmptyRow`。
- `ExcelImportFailureOptions.MaxBytes` 迁移为 `MaxSerializedBytes`。
- `ExcelCellValue` 构造签名增加 date-system 标识参数。

### Core

- 新增 `Bing.Offices.Dates.ExcelDateOffsetPolicy`，支持显式 offset 或配置固定 offset。
- 新增 public execution-detail `BingOfficesExceptionDispatcher`、`ObservedKey`、`ObserverFailureKey`，实际实现仍集中在 Core，供 Core/NPOI 复用。
- `ExcelDateAttribute` 新增 `OffsetPolicy`、`OffsetMinutes`。
- `DateTimeExcelValidationRule` 新增 `TryParseValue` 和 `TryParseWorkbookDate`，作为 NPOI 复用统一 parser 的窄边界。
- Core 与 NPOI 生产程序集之间不再使用 `InternalsVisibleTo`。

### NPOI

- 新增 Provider User API：`CellExtensions`、`RowExtensions`、`SheetExtensions`、`WorkbookExtensions`、`CellStyleExtensions`、`FontExtensions`。
- 公开扩展覆盖 cell value/conditional formatting/merge、row 创建、sheet 插入/合并/图片、workbook format/sheet、style/font mutation；内部 `InternalExtensions`、`PictureTypeResolver` 未公开。
- `AddNpoi` 迁移为 `AddBingOfficesNpoi`，旧入口未保留。

## 兼容性与批准

- 上述新增和删除属于本任务计划中的 Major breaking/API 收敛范围，但仍需要维护者对成员级 candidate diff 进行明确批准。
- 未修改正式 baseline 文件和 `PublicApiContractTest` 中的 formal hash。
- Unit 的唯一失败是 formal hash 不匹配，不能解释为行为测试失败，也不能通过修改 hash 断言消除。
- `ExcelDateOffsetPolicy`、dispatcher、日期 façade 被分类为 `Execution detail`/配置 User API 的治理项；`EditorBrowsable(Never)` 不等于 internal，批准前仍需明确其长期 public 责任。

## 验证

- 三个 candidate TFM 快照可生成，候选哈希一致。
- NPOI exact public type/member contract 通过。
- 生产程序集无生产 IVT 通过。
- 当前状态：`BLOCKED / formal API approval pending`。
