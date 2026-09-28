# 生产符号到测试追溯

| 生产符号或行为 | 关键验证 | 测试项目与方法 |
| --- | --- | --- |
| `Bing.Offices.EntityProviderContracts` 包 | 包含 net6.0/net8.0 合同 DLL 与 XML 文档，nuspec 仅依赖 Abstractions；第三方消费者不链接源码，双 TFM 运行六项合同 | `tests/Bing.Offices.ThirdPartyProvider.Consumer/Program.cs`（`VerifyCore`、`VerifyAttributeAndNamedAnchors`、`VerifyNamedListFixedCellCollision`、`VerifyPageSubtotal`、`VerifyGroupSubtotal`、`VerifyAsync`）；`Bing.Offices.ProviderContract.Tests.Contracts.EntityLayoutReusableContractTest`（双 TFM 各 13 项）；CI `Pack Entity Provider contracts` / `Third-party public-only Provider consumer` |
| `EntityLayoutAnalyzer` / `BOE001`-`BOE004` | 命名实参重排、继承属性、目标类型推断内联定义、条件表达式保守跳过；本地 NuGet 包实际加载 | `Bing.Offices.Analyzers.Tests.EntityLayoutAnalyzerTest`（net6.0/net8.0 各 8 项）；`tests/Bing.Offices.Analyzers.Consumer/Order.cs`（正常构建及 `BOE_NEGATIVE` 预期失败） |
| `ExcelEntityLayoutBuilder.Build` 命名列表冲突延迟校验 | 未解析的命名地址不会按占位 `A1` 误判；Provider 解析真实锚点后拒绝重叠 | `Bing.Offices.ProviderContract.Tests.Contracts.EntityLayoutReusableContractTest.AttributeAndNamedAnchors_ShouldSatisfyReusableContract`; `NamedListFixedCellCollision_ShouldSatisfyReusableContract`（NPOI/ClosedXML，net6.0/net8.0）；`Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_NamedListFixedCell_ShouldResolveActualAddress`（HSSF/XSSF） |
| `ExcelEntityCellAttribute` / `ExcelEntity.LayoutFromAttributes` | 属性固定单元格导出、导入和只读属性结构化错误 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_AttributesGroupsAndFooter_ShouldRoundTrip`; `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_ShouldRejectReadOnlyPropertyOnImport`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_AttributesGroupsAndFooter_ShouldRoundTrip`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_ShouldRejectReadOnlyPropertyOnImport` |
| `ExcelEntityLayoutBuilder.CellsFromAttributes` | 稳定属性顺序、地址解析、与 Fluent 配置混用 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_AttributesGroupsAndFooter_ShouldRoundTrip` |
| `DynamicColumnGroup` / `IExcelEntityDynamicColumnGroup` | 商品、产品等多个字典分组导出与按组回写 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_ShouldRoundTripDynamicGroupsAndFooter`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_AttributesGroupsAndFooter_ShouldRoundTrip` |
| 三字典商品档案布局 | 商品、产品、扩展字段顺序、空字典和未知键策略 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_ProductArchive_ShouldRoundTripThreeDynamicGroups`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_ProductArchive_ShouldRoundTripThreeDynamicGroups` |
| `UnknownDynamicValues` | 未知动态值策略进入导入计划 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_ShouldRoundTripDynamicGroupsAndFooter` |
| `ExcelDynamicColumnDefinition.ValidationRules` | 显式动态列组复用内置 range 校验并返回结构化坐标 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_ExplicitDynamicValidation_ShouldReportStructuredError`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_ExplicitDynamicValidation_ShouldReportStructuredError` |
| `ExcelEntityImportOptions.ValidationFailureMode` | 实体导入默认首错停止，Continue 收集当前行多个校验错误 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_ValidationFailureMode_ShouldControlRowErrors`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_ValidationFailureMode_ShouldControlRowErrors` |
| `CellNamed` / `ListRegionNamed` / `AnchorName` | 模板命名范围解析、局部名称优先、重复锚点、绝对结束地址和非法区域结构化拒绝 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_NamedAnchors_ShouldRoundTripTemplate`; `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_NamedAnchor_ShouldRejectInvalidTemplateRange`; `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_NamedAnchor_ShouldRejectDuplicateAndAbsoluteEnd`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_NamedAnchors_ShouldRoundTripTemplate`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_NamedAnchor_ShouldRejectMissingTemplateName`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_NamedAnchor_ShouldRejectDuplicateAndAbsoluteEnd` |
| `CalculatedColumn` / `ExcelEntityRowContext<TItem>` | 计算值、明细快照、零基索引、一基行列坐标、物理位置、数字格式、导入跳过和异常定位 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_CalculatedColumns_ShouldUseRowContextAndSkipOnImport`; `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_CalculatedColumnFailure_ShouldReportCoordinates`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_CalculatedColumns_ShouldUseRowContextAndSkipOnImport`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_CalculatedColumnFailure_ShouldReportCoordinates` |
| `GroupSubtotal` / `IExcelEntityGroupSubtotal` | 按相邻分组键写入组内聚合、支持最终 Footer，并在导入时跳过小计与间隔行 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_GroupSubtotal_ShouldRoundTripAndSkipSubtotalRows`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_GroupSubtotal_ShouldRoundTripAndSkipSubtotalRows` |
| Provider GroupSubtotal 合同 | 公开布局导出/导入、小计行隔离和 MiniExcel 不适用预期 | `Bing.Offices.ProviderContract.Tests.Contracts.EntityGroupSubtotalContractTest.EntityGroupSubtotals_ShouldMatchIndependentProfile`（net6.0、net8.0） |
| `PageBreak` / `ExcelEntityListRegion<TEntity>.PageBreakRows` | 每页明细行数写入水平分页符，表头不计入分页，导入边界和 Footer 语义保持不变；与 GroupSubtotal 组合时在构建阶段拒绝 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_PageBreak_ShouldWriteXssfRowBreaks`; `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_PageBreak_ShouldWriteHssfTemplateRowBreaks`; `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_PageBreak_ShouldRejectInvalidCombinations`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_PageBreak_ShouldWriteRowBreaks`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_PageBreak_ShouldRejectInvalidCombinations`; `tests/Bing.Offices.ThirdPartyProvider.Consumer/Program.cs` |
| `PageSubtotal` / `ExcelEntityListRegion<TEntity>.PageSubtotal` | 中间分页按页快照写入小计，最终 Footer 收尾，导入跳过小计与间隔；缺少 PageBreak、marker 缺失或位置错误时结构化拒绝 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_PageSubtotal_ShouldRoundTripAndBreakAfterSubtotal`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_PageSubtotal_ShouldRoundTripAndBreakAfterSubtotal`; `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_PageBreak_ShouldRejectInvalidCombinations`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_PageBreak_ShouldRejectInvalidCombinations`; `tests/Bing.Offices.ThirdPartyProvider.Consumer/Program.cs` |
| `PageSubtotal` marker 失败边界 | marker 缺失或出现在预期分页边界前时返回 Plan 配置错误，不把模板尾部误当明细 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_PageSubtotalMarkerErrors_ShouldBeStructured`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_PageSubtotalMarkerErrors_ShouldBeStructured` |
| `FooterNamed` / `ExcelEntityListRegion<TEntity>.FooterAnchorName` | 最终尾部名称随空明细和变长明细移动；模板同 Sheet 的全局或局部名称重新定位；冲突、取消和调用方目标保持完整 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_FooterNamed_ShouldCreateNameAndRoundTrip`; `Npoi_EntityLayout_FooterNamed_ShouldRelocateTemplateName`; `Npoi_EntityLayout_FooterNamed_ShouldPreserveOutputOnConflictAndCancellation`; `Npoi_EntityLayout_FooterNamed_ShouldRejectDuplicateAnchorsAndClearOnFooter`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_FooterNamed_ShouldCreateNameAndRoundTrip`; `EntityLayout_FooterNamed_ShouldRelocateTemplateName`; `EntityLayout_FooterNamed_ShouldPreserveOutputOnConflictAndCancellation`; `EntityLayout_FooterNamed_ShouldRejectDuplicateAnchorsAndClearOnFooter`; `tests/Bing.Offices.ThirdPartyProvider.Consumer/Program.cs`（net6.0、net8.0） |
| `IExcelEntityExporter.ExportEntity*` / `IExcelEntityImporter.ImportEntity*` | 可移植 XLSX 合同验证动态字典、Footer 名称、分页/分组小计、异步取消与调用方流所有权 | `Bing.Offices.ProviderContract.Tests.Contracts.EntityLayoutReusableContractTest.Core_ShouldSatisfyReusableContract`; `PageSubtotal_ShouldSatisfyReusableContract`; `GroupSubtotal_ShouldSatisfyReusableContract`; `Async_ShouldSatisfyReusableContract`（NPOI、ClosedXML，net6.0/net8.0）；`tests/Bing.Offices.ThirdPartyProvider.Consumer/Program.cs`（net6.0/net8.0） |
| Provider 分页小计合同 | 独立档案验证 NPOI/ClosedXML 往返和 MiniExcel 不适用预期 | `Bing.Offices.ProviderContract.Tests.Contracts.EntityPageSubtotalContractTest.EntityPageSubtotals_ShouldMatchIndependentProfile`（net6.0、net8.0） |
| `ExcelEntityListFooterBuilder<TItem>` | marker、GapRows、聚合快照、数字格式、样式和相对合并 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_ShouldRoundTripDynamicGroupsAndFooter`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_AttributesGroupsAndFooter_ShouldRoundTrip` |
| 无表头赠品单 Footer | 空明细、无表头、marker 和聚合单元格 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_GiftOrderWithoutHeaderAndEmptyDetails_ShouldRoundTrip`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_GiftOrderWithoutHeaderAndEmptyDetails_ShouldRoundTrip` |
| 多 Sheet 盘点单布局 | 固定汇总单元格和多个列表区域读回 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_Inventory_ShouldRoundTripMultipleSheetsAndRegions`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_Inventory_ShouldRoundTripMultipleSheetsAndRegions` |
| Footer marker 导入边界 | marker 缺失、重复、GapRows 前非法位置 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_FooterMarkerErrors_ShouldBeStructured`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_FooterMarkerErrors_ShouldBeStructured`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityTemplate_FooterGapWithTemplateContent_ShouldNotImportGapAsDetail` |
| 动态字典创建契约 | 空接口字典自动创建、只读属性、无参构造缺失和不兼容接口均为 Plan 错误 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_DynamicDictionaryCreationErrors_ShouldBeStructured`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_DynamicDictionaryCreationErrors_ShouldBeStructured` |
| Footer 物理行列边界 | HSSF/XSSF/ClosedXML 末行末列成功与越界前置失败，目标流保持空 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_FooterPhysicalBounds_ShouldBePreflightedForHssfAndXssf`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_FooterPhysicalBounds_ShouldBePreflighted` |
| ClosedXML Entity Layout 唯一性与资源上限 | 固定列跨行重复值、动态列重复值、唯一跟踪回滚和资源限制路径 | `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_ValidationMatrix_ShouldReportCompleteCoordinates`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_UniqueResourceLimit_ShouldReportStructuredError`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_UniqueValue_ShouldRollbackWhenAnotherColumnFails`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_DynamicUniqueColumn_ShouldReportDuplicateValue` |
| NPOI Entity Layout executor | XLSX/XSSF 成功往返、动态列、Footer、合并和异步职责 | `Bing.Offices.Npoi.Tests` EntityLayout 筛选（最终运行结果记录在 execution.md） |
| NPOI HSSF template Entity Layout | HSSF 动态分组、Footer 和导入边界 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.EntityTemplate_Hssf_ShouldRoundTripDynamicGroupsAndFooter` |
| ClosedXML Entity Layout executor | XLSX DOM 动态分组、Footer、样式、模板合并与回读 | `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_AttributesGroupsAndFooter_ShouldRoundTrip`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityTemplate_ShouldCreateVariableFooterMerge`; EntityLayout 筛选（最终运行结果记录在 execution.md） |
| Layout conflict and bounds validation | 非法地址、重叠列表区域、动态列冲突、Footer 越界和最大行列边界 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.EntityLayout_ShouldRejectInvalidAddressesAndOverlappingRegions`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_DynamicColumns_ShouldRejectCollisionAndOverflowBeforeWriting`; `Bing.Offices.ClosedXml.Tests.ClosedXmlProviderTest.EntityLayout_ShouldRejectFooterOutsideDeclaredBounds`; Provider-specific physical-bound tests above |
| Public package consumer | 只引用包公开 API，执行属性布局、两个动态组、分组小计和 Footer | `tests/Bing.Offices.ThirdPartyProvider.Consumer/Program.cs`；CI `consumer-net6` / `consumer-net8` 流程 |
| Documentation migration surface | Markdown C# fence 编译并执行，包括 NPOI 单据迁移示例 | `Bing.Offices.Docs.Tests.DocsConsumerTest.DocumentationFences_FromMarkdown_ShouldCompileAndExecuteIndividually` |

## Continuation 10：跨页签字区

| 最终生产符号 | 测试项目 | 测试方法与关键行为 |
|---|---|---|
| NpoiEntityLayoutSupport.ValidatePagedFooterBounds | Bing.Offices.Npoi.Tests | PageSignatures_ShouldRejectShiftedDetailOverflow：HSSF/XSSF 声明及物理边界，Plan 错误、原目标字节完整 |
| ClosedXmlEntityLayoutExecutor.ValidatePagedFooterBounds | Bing.Offices.ClosedXml.Tests | PageSignatures_ShouldRejectShiftedDetailOverflow：声明及物理边界，Plan 错误、原目标字节完整 |
| NpoiEntityExportExecutor.WritePagedTypedRegion / WriteFooter | Bing.Offices.Npoi.Tests | PageSignatures_ShouldRoundTripTemplate：HSSF/XSSF，空/整页/多页，有无表头、同步/异步、完整内容、合并/样式/行高/列宽/分页符和导入 |
| ClosedXmlEntityLayoutExecutor.WritePagedRegion / WriteFooter | Bing.Offices.ClosedXml.Tests | PageSignatures_ShouldRoundTripTemplate：同上 XLSX |
| NpoiEntityLayoutSupport.FindFooterSkipSpans | Bing.Offices.Npoi.Tests | Npoi_EntityLayout_PageSubtotalMarkerErrors_ShouldBeStructured：缺失、错位及无后续明细的终末标记 |
| ClosedXmlEntityLayoutExecutor.FindFooterSkipSpans | Bing.Offices.ClosedXml.Tests | EntityLayout_PageSubtotalMarkerErrors_ShouldBeStructured：同上 |

## Continuation 16：显式 Footer 公式

| 最终生产符号 | 测试项目 | 测试方法与关键行为 |
|---|---|---|
| `ExcelEntityListFooterBuilder<TItem>.Formula`、`ExcelEntityFooterFormulaValue.FormulaA1` | Bing.Offices.Npoi.Tests / Bing.Offices.ClosedXml.Tests | `Npoi_EntityLayout_FooterFormula_ShouldRejectInvalidFormulaAtBuild`、`EntityLayout_FooterFormula_ShouldRejectInvalidFormulaAtBuild`：缺少等号或公式体在构建阶段拒绝；包消费者 `Program.VerifyEntityLayout` 从 NuGet 包编译调用 |
| `NpoiEntityExportExecutor.WriteFooter` | Bing.Offices.Npoi.Tests | `Npoi_EntityLayout_FooterFormula_ShouldRoundTripForHssfAndXssf`：HSSF/XSSF 真实公式、数字格式、空/非空明细及导入边界；`Npoi_EntityLayout_FooterFormula_ShouldPreserveOutputOnInvalidFormula`：无效公式保留调用方流 |
| `ClosedXmlEntityLayoutExecutor.WriteFooter` | Bing.Offices.ClosedXml.Tests | `EntityLayout_FooterFormula_ShouldRoundTripAndSkipFooter`：XLSX 真实公式、数字格式、空/非空明细及导入边界 |

## Continuation 17：连续明细求和公式

| 最终生产符号 | 测试项目 | 测试方法与关键行为 |
|---|---|---|
| `ExcelEntityListFooterBuilder<TItem>.FormulaSumContiguousRowsAbove` | Bing.Offices.Npoi.Tests / Bing.Offices.ClosedXml.Tests | `Npoi_EntityLayout_FormulaSumContiguousRowsAbove_ShouldRoundTrip`、`EntityLayout_FormulaSumContiguousRowsAbove_ShouldRoundTrip`：HSSF/XSSF/XLSX、空/非空明细、`GapRows` 顺序和构建后快照、相对行、格式与导入；`Npoi_EntityLayout_FormulaSumContiguousRowsAbove_ShouldPreserveOutputOnPlanFailure`、`EntityLayout_FormulaSumContiguousRowsAbove_ShouldPreserveOutputOnPlanFailure`：目标流保持完整；包消费者 `Program.VerifyEntityLayout` 双 TFM 公开调用 |
| `ExcelEntityLayoutBuilder<TEntity>.CreateListRegion` | Bing.Offices.Npoi.Tests / Bing.Offices.ClosedXml.Tests | `Npoi_EntityLayout_FormulaSumContiguousRowsAbove_ShouldRejectInvalidColumnsAndFinalFooterConflicts`、`EntityLayout_FormulaSumContiguousRowsAbove_ShouldRejectInvalidColumnsAndFinalFooterConflicts`：非法列及最终 Footer 与中间小计冲突；`Npoi_EntityLayout_FormulaSumContiguousRowsAbove_ShouldWorkForPageSubtotal`、`ShouldWorkForGroupSubtotal`、`EntityLayout_FormulaSumContiguousRowsAbove_ShouldWorkForPageSubtotal`、`ShouldWorkForGroupSubtotal`：中间小计真实公式与导入边界 |

## 公式 Provider 合同阶段

| 最终生产符号与合同入口 | 测试项目 | 测试方法与关键行为 |
|---|---|---|
| `ExcelEntityListFooterBuilder<TItem>.Formula`、`FormulaSumContiguousRowsAbove` → `EntityLayoutProviderContractSuite.VerifyFooterFormulas` | Bing.Offices.ProviderContract.Tests / Bing.Offices.ThirdPartyProvider.Consumer | `EntityLayoutReusableContractTest.FooterFormulas_ShouldSatisfyReusableContract`：NPOI、ClosedXML 在 net6.0/net8.0 验证原生公式、连续范围、空明细、marker 与冲突；`Program.VerifyEntityLayout`：双 TFM 从合同 NuGet 包调用 |

## Continuation 19：跨小计明细总计公式

| 最终生产符号 | 测试项目 | 测试方法与关键行为 |
|---|---|---|
| `ExcelEntityListFooterBuilder<TItem>.FormulaSumDetailRowsAbove`、`ExcelEntityFooterDetailSumValue.ToFormulaA1` | Bing.Offices.ProviderContract.Tests | `DetailSumFormula_ShouldValidateAndPartitionDetailRanges`：来源列标准化、空明细、非连续区间、非法行段、SUM 参数分段与公式长度上限 |
| `NpoiEntityExportExecutor.WriteFooter`、明细行段收集 | Bing.Offices.Npoi.Tests | `Npoi_EntityLayout_FormulaSumDetailRowsAbove_ShouldSkipPageSubtotals`、`ShouldSkipGroupSubtotals`：HSSF/XSSF 的分页与分组、多行小计、间隔行、空明细、数字格式、marker 导入边界；`ShouldRejectIntermediateSubtotalAtBuild`：中间小计禁用该方法 |
| `ClosedXmlEntityLayoutExecutor.WriteFooter`、明细行段收集 | Bing.Offices.ClosedXml.Tests | `EntityLayout_FormulaSumDetailRowsAbove_ShouldSkipPageSubtotals`、`ShouldSkipGroupSubtotals`、`ShouldRejectIntermediateSubtotalAtBuild`：同等 XLSX 合同 |
| `EntityLayoutProviderContractSuite.VerifyFooterFormulas` | Bing.Offices.ProviderContract.Tests / Bing.Offices.ThirdPartyProvider.Consumer | `FooterFormulas_ShouldSatisfyReusableContract`：双 TFM 公开合同验证跨分页小计时仅汇总明细行；`Program.VerifyEntityLayout`：双 TFM 从 NuGet 包调用公开方法 |

## Continuation 20：静态区域冲突诊断

| 最终生产符号 | 测试项目 | 测试方法与关键行为 |
|---|---|---|
| `EntityLayoutAnalyzer.AnalyzeLayoutAreas`、`OverlappingAreaId` (`BOE005`) | Bing.Offices.Analyzers.Tests | `OverlappingMerges_ShouldReportOnce`、`BoundedList_ShouldReportKnownAreaConflicts`、`BoundedLists_ShouldReportOverlap`：常量合并、固定单元格和有界列表的交叉占位；`RepeatedEnd_ShouldUseLastAddress`：最终生效的列表边界 |
| `EntityLayoutAnalyzer.CollectAttributedCells`、`Conflicts` | Bing.Offices.Analyzers.Tests | `AttributeCellAndBoundedList_ShouldReportConflict`、`ExplicitAttributeCells_ShouldReportConflict`、`AttributeAndFluentCell_ShouldReportSharedAddress`、`MergedCells_ShouldReportNormalizedAnchorCollision`：属性式、Fluent 和合并锚点冲突；`ValidAndUnknownAreas_ShouldNotReport`：合法合并、不同工作表和未知边界不误报 |
| `BOE005` NuGet 诊断 | Bing.Offices.Analyzers.Consumer | `BOE_AREA_NEGATIVE` 编译负例报告 `error BOE005`；正常消费者编译通过，原 `BOE_NEGATIVE` 继续报告 `error BOE001` |

## Continuation 21：报表级导出策略

| 最终生产符号 | 测试项目 | 测试方法与关键行为 |
|---|---|---|
| `ExcelWorkbookExportMode`、`ExcelWorkbookExportStrategy.Export`、`ExportAsync` | Bing.Offices.Tests.Integration | `CompleteWorkbook_ShouldWriteReadableWorkbook`：NPOI XLS/XLSX 与 ClosedXML XLSX 同步/异步完整数据；`CompleteWorkbook_ShouldAcceptTemplate`：完整模式保留 NPOI/ClosedXML 模板能力；`ForwardStreaming_ShouldWriteReadableWorkbook`：SpreadCheetah 同步/异步真实 XLSX；`UnsupportedSelection_ShouldLeaveDestinationEmpty`、`MissingOrIncompatibleProvider_ShouldFailBeforeOutput`：格式、模板、缺失模式和未知模式预检后无输出 |
| `ExcelWorkbookExportStrategy.ExportToFile`、`ExportToFileAsync` | Bing.Offices.Tests.Integration | `FileExport_ShouldCommitSelectedWorkbook`：两种模式同步/异步真实文件替换；`PreCancelledFileExport_ShouldPreserveTarget`：预取消不改变已有目标 |
| `ExcelWorkbookExportStrategy` 包消费者入口 | Bing.Offices.ThirdPartyProvider.Consumer | `Program.VerifyWorkbookExportStrategy`：net6.0/net8.0 从本地 NuGet 包执行完整与前向模式，并验证模板预检及流所有权 |

## Continuation 22：第三方 Provider 模板

本阶段未修改生产程序集。以下为模板示例符号到独立测试的映射。

| 模板符号或行为 | 测试项目 | 测试方法与关键行为 |
|---|---|---|
| `StarterExcelExporter.Export`、`ExportAsync` | Starter.Provider.Tests | `Export_ShouldWriteAllSheets`、`ExportAsync_ShouldUseDestinationAsyncIo`：net6.0/net8.0 真实 XLSX 多 Sheet 完整内容、异步目标写入与流所有权 |
| `StarterExcelExporter` 能力声明与请求预检 | Starter.Provider.Tests | `Capabilities_ShouldMatchImplementedBoundary`、`UnsupportedRequest_ShouldLeaveDestinationUntouched`、`EntityLayout_ShouldRejectBeforeOutput`：已实现能力与不支持请求的输出前拒绝 |
| `AddStarterExcelProvider` 与 `IFileExportCommitter` | Starter.Provider.Tests | `AddStarterExcelProvider_ShouldPreserveHostServices`、`ExportToFile_ShouldReplaceExistingWorkbook`、`PreCancelledExportToFile_ShouldPreserveExistingBytes`：宿主注入、真实文件替换和预取消保护 |
| 资源限制与异常观察 | Starter.Provider.Tests | `FailedBuild_ShouldPreserveDestinationAndObserveOnce`：行数或构建失败保留目标并且仅观察一次 |

## Continuation 23：旧版 API 迁移诊断

| 最终生产符号 | 测试项目 | 测试方法与关键行为 |
|---|---|---|
| `LegacyApiMigrationAnalyzer` / `BOM001` | Bing.Offices.Analyzers.Tests / Bing.Offices.Analyzers.Consumer | `LegacyExportTypes_ShouldReportMigrationWarnings`：已绑定旧导出服务和选项各报一次；`UnrelatedNames_ShouldNotReport`：不同命名空间同名类型不误报；`BOM_MIGRATION` 本地包消费者输出两条 `BOM001` |
| `LegacyApiMigrationAnalyzer` / `BOM002` | Bing.Offices.Analyzers.Tests / Bing.Offices.Analyzers.Consumer | `LegacyDynamicMarker_ShouldWarnWithoutFlaggingPreservedAttributes`：旧标记报 Warning，保留的 `DynamicColumn`/`ColumnName` 不告警；`BOM_MIGRATION` 本地包消费者输出 `BOM002` |
