# Public API approval — failure workbook committers

- approvedBy: user (explicit implementation-plan approval)
- approvedAt: 2026-09-27 (Asia/Shanghai)
- taskId: BING-OFFICES-PUBLIC-PROVIDER-BOUNDARY-20260927-001
- Evidence: 用户明确要求实施包含下述两个六参数公开构造重载的完整计划。

## Approved member additions (net6.0 / net8.0)

1. `Bing.Offices.Imports.NpoiExcelImporter(IEnumerable<IExcelValidationRule> validationRules, IEnumerable<IExcelValueConverter> valueConverters, IEnumerable<INamedExcelValidationRule> namedValidationRules, IExcelMappingPlanFactory mappingPlanFactory, IEnumerable<IBingOfficesExceptionObserver> exceptionObservers, IFileExportCommitter fileExportCommitter)`
2. `Bing.Offices.ClosedXml.Imports.ClosedXmlExcelImporter(IEnumerable<IExcelValidationRule> validationRules, IEnumerable<IExcelValueConverter> valueConverters, IEnumerable<INamedExcelValidationRule> namedValidationRules, IExcelMappingPlanFactory mappingPlanFactory, IEnumerable<IBingOfficesExceptionObserver> exceptionObservers, IFileExportCommitter fileExportCommitter)`

分类：Provider User API。六个参数均必传；提交器 null 使用默认实现。
原五参数构造函数及默认值不变；无公共成员删除或签名替换。

## Compatibility and migration

现有源码和五参数构造调用无需修改。需要自定义失败工作簿文件提交时，使用新六参数构造或在 AddBingOfficesNpoi / AddBingOfficesClosedXml 前注册 IFileExportCommitter。旧 Provider 二进制曾使用 Core 内部 AtomicFileCommitter，必须与本次 Core 配套升级；源码公共 API 兼容不等同于旧生产友元二进制兼容。

Core 的生产友元移除属于本次明确批准的程序集边界变更；内部原子提交工具和测试文件系统不公开。更新快照前保存逐成员增删差异；不得将无关差异自动批准。

## Earlier approvals

既有公共 API 不变，历史审批继续见 [上一版 API 审批](../BING-OFFICES-ADVANCED-EXCEL-ROADMAP-20260926-001/api-approval.md)。本文件仅批准上述两个增量和本次友元移除。
