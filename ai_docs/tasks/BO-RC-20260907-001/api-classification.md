# API 分类台账

本轮以 `PublicApiContractTest` 的完整 exported type ledger 为主；历史精确数量为 Abstractions/Core/NPOI 合计 165 个类型（141 User API、9 Provider User API、15 Provider SPI）。新增/变更项如下：

| API | Assembly | Visibility | Category | Decision | Reason |
| --- | --- | --- | --- | --- | --- |
| `NpoiExcelImporter` | Npoi | public sealed | Provider User API | 保留 | 构造器稳定，允许不依赖 DI 直接实例化 |
| `NpoiExcelExporter` | Npoi | public sealed | Provider User API | 保留 | 构造器稳定，允许不依赖 DI 直接实例化 |
| `CellExtensions`/`CellStyleExtensions`/`FontExtensions` | Npoi | public static | Provider User API | 保留 | 直接操作 NPOI 类型的正式用户扩展 |
| `RowExtensions`/`SheetExtensions`/`WorkbookExtensions` | Npoi | public static | Provider User API | 保留 | 直接操作 NPOI 类型的正式用户扩展 |
| `ExcelNpoiServiceCollectionExtensions` | Npoi | public static | User API | 保留 | DI 注册入口 |
| `IExcelImporter`/`IExcelExporter`/`ICsvImporter`/`ICsvExporter` | Abstractions | public interface | User API | 增加 Async 成员 | 同步/异步双合同，Breaking 允许且需迁移 |
| `IFileExportCommitter` | Abstractions | public interface + Never | Provider SPI | 增加 CommitAsync | 文件提交边界由 exporter 使用，第三方可替换 |
| `NpoiStreamCopier`、Failure writer、plan/executor/cache | Npoi | internal | Internal | 保持隐藏 | 实现细节，不表达业务需求 |

生产程序集仍只保留测试友元 `InternalsVisibleTo`；没有新增生产程序集间 IVT。

## Candidate 删除成员台账（待维护者审批）

以下删除项来自 `artifacts/api-compare-final/api-diff.json`，不能由执行器自行视为已批准。它们与本轮 Async/namespace 变更没有直接调用链关系，维护者需要在更新正式 baseline 前明确确认是否属于既有基线漂移，并决定是否恢复或提供迁移版本说明。

| 成员 | Assembly | 当前处置 | 迁移/审批要求 |
| --- | --- | --- | --- |
| `ExcelMappingDiagnostic` 类型 | Abstractions | `PENDING_APPROVAL` | 确认删除是否有既有替代诊断模型；若保留删除，发布说明必须列出 Breaking Change |
| `ExcelMappingDiagnostic` 构造函数及 `Code`/`Message`/`Path` 属性 | Abstractions | `PENDING_APPROVAL` | 与类型删除一起审批，不能只批准 hash |
| `ExcelMappingConfigurationLoader.FromJsonDocument(string, out IReadOnlyList<ExcelMappingDiagnostic>)` | Core | `PENDING_APPROVAL` | 提供无 diagnostics overload 或新的诊断返回合同迁移说明 |
| `ExcelMappingConfigurationLoader.FromXmlDocument(string, out IReadOnlyList<ExcelMappingDiagnostic>)` | Core | `PENDING_APPROVAL` | 同上；Package Consumer/Docs 需覆盖迁移后的调用方式 |
