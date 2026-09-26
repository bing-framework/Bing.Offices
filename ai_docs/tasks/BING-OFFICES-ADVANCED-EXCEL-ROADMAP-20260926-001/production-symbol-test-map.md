# 本轮生产符号 → 职责级测试

范围为 F01–F06。下方历史符号清单不代表 Aspose/公式等在本轮验收。

| 生产符号 | 关键行为 | 测试项目 / 方法 |
| --- | --- | --- |
| `ExcelSheetExportBuilder.Image/DataValidation`、`ExcelSheetExportRequest.Images/DataValidations` | 请求快照与无效定义 | ProviderContract / `SheetContentContractTest.Builder_ShouldSnapshotImageAndListAndRejectInvalidDefinitions` |
| `ExcelSheetContent.Validate` | 格式边界、输出前拒绝 | ProviderContract / `Export_OutOfRange_ShouldFailBeforeEnumerationOrOutput` |
| `NpoiSheetContentWriter.Apply/Write` | XLS/XLSX 图片、32 比较、列表引号修正 | ProviderContract / `Export_ShouldPreserveImageAndEveryValidationOperator`, `Export_JpegAcrossSheets_ShouldPreserveContents` |
| `ClosedXmlSheetContentWriter.Apply` | 图片、原生校验完整结构 | ProviderContract / 同上两方法（ClosedXML Profile） |
| `SpreadCheetahSheetContentWriter` | 96 DPI 锚点、JPEG 部件、图片暂存 | ProviderContract / 同上两方法及 `SpreadImages_FailedEnumeration_ShouldCleanStagingAndPreserveFile` |
| `MiniExcelExcelExporter` 预检 | 新内容 Unsupported、零输出 | ProviderContract / 同上内容测试（MiniExcel Profile） |
| `SpreadCheetahStreamingExcelExporter` 列绑定 | ValueMap、Converter、动态列布局 | SpreadCheetah.Tests / `ExportBatchesAsync_ShouldApplyValueMapNamedConverterAndDynamicPhysicalLayout` |
| 同上映射预检 | 配置失败零枚举、完整异常元数据 | SpreadCheetah.Tests / `ExportBatchesAsync_ShouldRejectInvalidMappingBeforeEnumeratingOrWriting`, `ExportBatchesAsync_WhenNamedConverterIsMissing_ShouldReturnStructuredConfigurationErrorBeforeEnumerating` |
| 同上映射异常边界 | 固定/动态转换、Getter、Formatter 的位置与错误分类 | SpreadCheetah.Tests / `ExportBatches_WhenValueConverterFails_ShouldReturnStructuredExportError`, `ExportBatches_WhenDynamicGetterFails_ShouldReturnStructuredExportError`, `ExportBatches_WhenDynamicValueCannotConvert_ShouldReturnStructuredExportError`, `ExportBatches_WhenPropertyGetterFails_ShouldReturnStructuredExportError`, `ExportBatches_WhenDynamicConverterFails_ShouldReturnStructuredExportError`, `ExportBatches_WhenFormatterFails_ShouldReturnStructuredExportError` |
| 同上映射取消边界 | 同步/异步取消对象原样传播 | SpreadCheetah.Tests / `ExportBatches_WhenPropertyGetterCancels_ShouldPreserveOriginalException` |
| 同上批次循环 | 0/1/999/1000/1001、单次枚举、实际 IO 背压 | SpreadCheetah.Tests / `ExportBatchesAsync_ShouldHandleBoundaryCountsWithSingleEnumeration`, `ExportBatchesAsync_WithLargeSourceAndAsyncWriteGate_ShouldPauseBeforeSourceCompletes` |
| 同上流与文件入口 | 中途取消、故障、所有权、旧目标/临时文件 | SpreadCheetah.Tests / `ExportBatchesAsync_WhenDestinationCancelsDuringWrite_ShouldPropagateAndLeaveStreamOpen`, `ExportBatchesAsync_WhenDestinationFails_ShouldPropagateWriteFailureAndLeaveStreamOpen`, `ExportBatchesToFileAsync_WhenCanceledMidWrite_ShouldPreserveSentinelAndCleanTemporaryFile`, `ExportBatchesToFileAsync_WhenEnumeratorFails_ShouldPreserveSentinelAndCleanTemporaryFile` |
| `ExcelReportPreflight`、`NpoiReportWriter` | 高级规则、表格/名称/冻结/打印、XLS 拒绝 | Npoi.Tests / `NpoiReportContractTest.Export_CommonReportDefinitions_ShouldWriteWorkbookStructure`, `Export_AdvancedConditionalFormats_ShouldWriteExpectedXlsxRules`, `Export_XlsAdvancedConditionalFormat_ShouldFailBeforeEnumeration`, `Export_XlsTable_ShouldFailBeforeEnumeration`, `Export_PrintScaleConflict_ShouldFailBeforeEnumeration`, `Export_DuplicateWorkbookName_ShouldFailBeforeEnumeration`, `Export_OutOfBoundsRange_ShouldFailBeforeEnumeration` |
| `ExcelReportPreflight`、`ClosedXmlReportWriter` | 等价高级规则与预检 | ClosedXml.Tests / `ClosedXmlReportContractTest.Export_CommonReportDefinitions_ShouldWriteWorkbookStructure`, `Export_AdvancedConditionalFormats_ShouldWriteExpectedXlsxRules`, `Export_InvalidReportDefinition_ShouldFailBeforeEnumeration` |
| 公共新增成员 | 双 TFM、增量兼容 | Tests.Integration / `PublicApiContractTest`；ApiSnapshotVerifier |
| `ExcelSheetContent.DateSerial` 与各 Provider 日期规则 | 1900 早期边界、1904 模板 | ProviderContract / `Export_DateValidation_ShouldPreserveEarly1900Serial`, `Export_DateValidation_ShouldRespect1904Template` |
| `ExcelReportPreflight` 与颜色 writer | 倒序、打印/冻结边界及缺省色一致性 | ProviderContract / `Export_ReportBoundary_ShouldRejectBeforeEnumeration`, `Export_ConditionalDefaults_ShouldUseEquivalentColors` |
| SpreadCheetah 报表准备阶段 | 后续 Sheet 静态错误零输出，重复名称/冲突筛选/越界；同范围表格筛选合并 | SpreadCheetah.Tests / `ExportBatchesAsync_ShouldPreflightLaterSheetBeforeAnyOutput`, `ExportBatchesAsync_ShouldMergeSameRangeTableAndFilter` |
| SpreadCheetah 表格行数边界 | 不预枚举，数据与范围不符不静默扩缩 | SpreadCheetah.Tests / `ExportBatches_TableRangeMismatch_ShouldFailWithoutSecondEnumeration` |
| `NpoiReportWriter` HSSF 分支 | 普通条件规则的调色板字体与实心填充 | Npoi.Tests / `Export_XlsBasicConditionalFormat_ShouldRoundTripPaletteColors` |
| NPOI/ClosedXML 打印 writer | Letter、10/400 缩放、带 `$` 重复标题及非法缩放预检 | 各 Provider ReportContract / `Export_PrintLayout_ShouldRoundTrip`, `Export_PrintScaleBelowProviderMinimum_ShouldFailBeforeEnumeration` |
| `NpoiReportWriter` 汇总行与共享预检 | 追加空汇总行，保留所有业务数据及原筛选区域；末行溢出提前拒绝 | Npoi.Tests / `Export_CommonReportDefinitions_ShouldWriteWorkbookStructure`；Npoi.Tests、ClosedXml.Tests / `Export_TableTotalsBeyondLastRow_ShouldFailBeforeEnumeration` |
| `NpoiReportWriter` 名称作用域 | XLS/XLSX 全局与两个 Sheet 的同名局部名称保存读回 | Npoi.Tests / `Export_SameNamedWorkbookAndSheetRanges_ShouldPreserveScopes` |
| SpreadCheetah 包发布资产 | 真实调用与独立读回完整行 | build/PackageConsumers/Consumer.Net6、Consumer.Net8 / `Program` |
| Docker/容量工具 | 固定运行时、非 root、只读源、4 GiB、9 容量单元 | `scripts/verify-docker.sh`；`build/StreamingProbe`；根 artifacts/open-provider-20260926 |

## 历史清单（非本轮新增验收）

| 生产符号 | 关键行为 | 直接测试 |
|---|---|---|
| `ExcelFormat` | 保留旧值并追加 Xlsm/Ods | `PublicApiContractTest` |
| `ExcelSheetExportBuilder` | 报表定义校验与不可变请求 | `ClosedXmlReportContractTest`, `NpoiReportContractTest` |
| `NpoiReportWriter` | XLSX 表格、筛选、冻结、名称、条件格式、打印 | `NpoiReportContractTest` |
| `ClosedXmlReportWriter` | XLSX 表格、筛选、冻结、名称、条件格式、打印 | `ClosedXmlReportContractTest` |
| `NpoiExcelFormulaProcessor` | 读取公式文本/缓存，拒绝重新计算 | `NpoiFormulaContractTest` |
| `ClosedXmlExcelFormulaProcessor` | 读取公式文本/缓存，拒绝重新计算 | `ClosedXmlFormulaContractTest` |
| `AsposeCellsEngine` | 无许可证预检、能力和密码诊断边界；真实渲染/转换由许可证 Golden 验证 | `AsposeCellsProviderContractTest` |
| `IExcelStreamingExporter` | 前向新建、批次、真实异步 IO、文件提交 | `SpreadCheetahStreamingExcelExporterTest` |
| `docker/verify/Dockerfile` | .NET 6/8、fontconfig、DejaVu/Noto CJK | `scripts/verify-docker.sh`, CI build/test jobs |
