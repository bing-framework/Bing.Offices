# 生产符号 → 职责测试

| 符号 | 行为 | 测试项目与方法 |
| --- | --- | --- |
| `ClosedXmlWorkbookValidationPipeline.TryResolveNamedRange` | 局部优先、跨 Sheet 名称与单格 | ClosedXml.Tests / `NamedList_ShouldResolveLocalGlobalSingleCellAndEscapedSheetNames` |
| `TryResolveDefinedName` / `TrySplitReference` | 仅单个绝对一维有界区域；拒绝动态、联合、相对等定义 | ClosedXml.Tests / `NamedList_UnsupportedDefinitions_ShouldReturnStructuredUnsupported` |
| `TryResolveNamedRange` | 无效局部定义不回退 | ClosedXml.Tests / `NamedList_InvalidLocalName_ShouldNotFallBackToGlobalName` |
| `TryResolveDefinedName` | 循环别名与缺失 Sheet 明确拒绝 | ClosedXml.Tests / `NamedList_CyclicAliasAndMissingSheet_ShouldRemainUnsupported` |
| `ValidateList` | 直接外部列表表达式 Unsupported；不等同于外部名称定义 | ClosedXml.Tests / `WorkbookValidationUnresolvableReferences_ShouldBeExplicitlyUnsupported` |
| `ValidateListRange` | 完整检查名称源，匹配前项不绕过后续公式 | ClosedXml.Tests / `NamedList_FormulaCellAfterMatch_ShouldRemainUnsupportedWithoutRecalculation` |
| `ValidateList` | 保留显式列表单项空格和逗号项修剪行为 | ClosedXml.Tests / `ExplicitList_ShouldPreserveExistingSingleAndCommaSeparatedSemantics` |
| `Validate` | IgnoreBlanks、预取消；范围循环逐项检查取消由源码核对 | ClosedXml.Tests / `NamedList_BlankAndCancellation_ShouldUseExistingPolicies` |
| `ClosedXmlExcelImporter.ImportAsync` | 预取消保留输入流 | ClosedXml.Tests / `ImportAsync_PreCanceled_ShouldKeepSourceOpen` |
| `ClosedXmlExcelImporter.Import/ImportAsync` → 原生列表校验 | 保存后真实导入、有效实体、源流所有权 | ClosedXml.Tests / `Import_NamedListValidValue_ShouldRoundTripAndKeepSourceOpen` |
| 同上 | 非法值即使 Report 也拒绝 | ClosedXml.Tests / `Import_NamedListInvalidValue_ShouldRejectRowUnderReportPolicy` |
| 同上 | Unsupported 在 Report 保留实体、Fail 拒绝 | ClosedXml.Tests / `Import_NamedListFormulaSource_ShouldRespectUnsupportedPolicy` |
| `scripts/verify-docker.sh` | 主环境版本与镜像身份真实采集 | `--metadata` exit 0；`artifacts/named-list-20260926/docker/docker-version.json`、`image-identity.json` |

既有 WorkbookValidationContractTest 覆盖未定义名称、直接范围、比较规则与配置校验顺序；本任务运行该合同的双 TFM 回归。公共 API 形状由 PublicApiContractTest 双 TFM 验证，无新增成员。
