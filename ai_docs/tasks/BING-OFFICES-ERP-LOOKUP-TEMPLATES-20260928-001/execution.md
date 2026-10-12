<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-ERP-LOOKUP-TEMPLATES-20260928-001
AI_EXECUTION_FINISHED_AT: 2026-09-29T01:41:14Z

# 实施执行报告

## 执行结论

已完成 Workbook 和 Entity 的隐藏文本字典、名称列表下拉，以及 NPOI/ClosedXML 的名称列表导入和 Entity 原生校验接入。现有构造、接口和 Fluent 调用保留；未提交、推送或发布。

## 任务信息

- 开始时间：2026-09-28T16:01:17.525Z。
- 执行依据：本任务 `plan.md` 和用户批准的实施计划。
- 外部业务库、版本号：未修改。

## 计划执行情况

- P1 公共契约：完成。字典项立即规范化并快照；列表来源互斥；Entity 可声明按 Sheet 的原生规则，旧导入选项维持仅配置校验。
- P2 Provider：完成。NPOI XLS/XLSX、ClosedXML XLSX 建立隐藏 Sheet 和工作簿名称，再应用下拉；NPOI 对齐名称解析、作用域优先、Unsupported 策略及单次导入缓存；Entity 固定单元格和明细行保留源坐标。MiniExcel 与 SpreadCheetah 对新请求预检拒绝。
- P3 验证与交付：完成核心行为、双 TFM、包消费者、API 审批、文档与常规回归。原有文件提交、取消、临时文件和流所有权测试随全量回归运行；本轮没有重跑无关性能矩阵。

## 已完成事项

- 增加 `LookupList`、`ListSourceName`、Entity `DataValidation` 与导入验证选项；增加隐藏于 IntelliSense 的字典 Provider SPI 和 `NamedListValidation` 能力位。
- 增加 XLS/HSSF、XLSX/XSSF 与 ClosedXML 的真实文件结构测试、混合内联/名称列表测试、Entity 导入往返和结构化错误测试。
- 将 NuGet 包消费者扩展为公开 Workbook/Entity 字典模板示例，覆盖 net6.0 和 net8.0。
- 更新导入校验说明、Provider 能力矩阵、迁移与审批记录、生产符号测试追溯表。

## 部分/未完成事项

- 本任务范围内无未完成的核心功能。填写区域保护、级联字典、分批导入扩展、分组分页组合与透视报表仍属于后续独立任务。
- 未新增独立的大批量字典性能基准；名称解析缓存由职责测试和代码路径核对，资源限制沿用现有测试。部署容量仍须按目标环境另行测量。

## 修改文件

- 公共契约：`src/Bing.Offices.Abstractions` 的 Exports、Entities、Providers 相关文件。
- Provider：`src/Bing.Offices.Npoi`、`src/Bing.Offices.ClosedXml` 的导入/导出路径，MiniExcel 与 SpreadCheetah 的预检路径。
- 测试与文档：两个 Provider 测试项目、双 TFM Consumer、公共 API 合同测试、`docs/excel` 和本任务文档。
- API 基线：`build/api-snapshot-baseline.json`，按成员差异核对后更新。

## API/数据/配置变化

net6.0 与 net8.0 的公共 API 差异一致：只在 `Bing.Offices.Abstractions` 新增 14 个成员，删除 0 个；详情见 `api-diff.json` 与 `api-approval.md`。字典值会去除首尾空白、拒绝空项并按 Ordinal 去重；旧显式列表行为保持不变。

## 测试结果

- NPOI 字典结构与混合列表：net6.0、net8.0 各 5/5；ClosedXML 对应测试各 2/2。
- 双 Provider 的 Entity 原生校验、名称列表规则和公共模型职责测试通过。
- 八个生产包 Release 打包及新包 Consumer：net6.0、net8.0 均输出 `package-consumer-ok`。
- 双 TFM 公共 API 合同：各 11/11；文档测试：19/19。
- `dotnet test Bing.Offices.sln -c Release --no-build --no-restore -m:1 --filter Category!=Large --verbosity quiet`：全部项目与两个目标框架通过。

## Build/Typecheck/Lint/Format

`dotnet build Bing.Offices.sln -c Release --no-restore --verbosity quiet` 通过，0 错误。严格 UTF-8、BOM、行尾和末尾换行检查通过；`git diff --check` 通过。构建仍有仓库既存的 Microsoft.Bcl.Memory net6.0 兼容性警告。

## 计划偏差

首次并行全量测试出现 NPOI net8.0 取消测试的瞬时失败；两例单独复跑及最终 `-m:1` 全量测试均通过。未据此修改生产取消逻辑。

## 基线问题

预存的未跟踪 `tests/Bing.Offices.EntityProviderContracts/Bing.Offices.EntityProviderContracts.xml` 未触碰。Git 对 `tests/Bing.Offices.ProfileFixtures/Bing.Offices.ProfileFixtures.xml` 提示下一次触碰时会把 CRLF 转 LF；该文件未修改。

## 已知问题

没有本任务已知的功能性失败。工作簿原生校验仅在显式选择相应导入模式时执行；动态公式、名称链、二维/联合/外部名称引用仍按 Unsupported 策略处理。

## 风险与回归关注点

大字典会增加工作簿大小和 DOM 内存；应按既有资源限制设置最大行数、单元格长度和输入大小。直接写入调用方 Stream 失败时不能回滚已写字节，文件目标仍使用原子提交。

## Reviewer 注意事项

重点核对名称局部优先且无效局部不回退、XLSX 内联列表 XML 修正只作用于显式列表、Entity 失败行不物化以及源坐标映射。成员审批及测试映射见本任务同目录文件。

## Git 状态

保留工作区改动，未执行 `git commit`、`git push`、发布或版本号修改。

## Review 修复记录

### Round 1

- Review 状态：NEEDS_FIX；Fix Scope：recommended。
- Review 文件：`ai_docs/tasks/BING-OFFICES-ERP-LOOKUP-TEMPLATES-20260928-001/review.md`；保留其独立审查结论，未修改。

#### FIX-001（HIGH / MUST_FIX）：COMPLETED

NPOI XLSX 内联列表 XML 修正改为按区域和原显式公式定位节点，并避免重复改写同一规则。`LookupListExportStructureTest.Export_OverlappingLists_ShouldPreserveNamedFormula` 覆盖两种添加顺序，名称公式保持完整。

#### FIX-002（HIGH / MUST_FIX）：COMPLETED

NPOI 和 ClosedXML 的 Workbook 导出均先建立业务 Sheet 与手工名称，再建立自动字典，写入数据后统一应用校验。`LookupListExportStructureTest.Export_NamedList_ShouldResolveLaterSheet` 覆盖 NPOI XLS/XLSX 与 ClosedXML XLSX 的前向引用。

#### FIX-003（MEDIUM / SHOULD_FIX）：COMPLETED

两个 Provider 在应用名称列表规则前校验局部优先、绝对有界一维来源和公式字典边界。`LookupListExportStructureTest.Export_TwoDimensionalNamedList_ShouldLeaveOutputEmpty` 验证非法来源在目标写入前失败；既有名称作用域、公式和 Unsupported 用例继续通过。

#### FIX-004（MEDIUM / SHOULD_FIX）：COMPLETED

ClosedXML Entity 明细先执行整行工作簿原生校验，再进行转换和配置物化。`EntityNativeValidationTest.EntityImport_InvalidNativeValue_ShouldSkipLaterConversionInSameRow` 确认原生失败不会附带后续类型转换错误；两个 Provider 的分页小计原始行坐标用例通过。

#### FIX-005（MEDIUM / SHOULD_FIX）：COMPLETED

NPOI 与 ClosedXML 的名称引用解析只在引号外检查联合、外部和公式语法。`NamedListValidationTest.NamedList_QuotedSheetNameWithSpecialCharacters_ShouldImport`、`NamedListValidationTest.NamedList_QuotedSheetNameWithSpecialCharacters_ShouldValidate` 覆盖逗号、括号、感叹号及转义单引号工作表名称。

#### FIX-006（MEDIUM / SHOULD_FIX）：COMPLETED

NPOI 将允许值改为 Ordinal `HashSet`，ClosedXML 按工作簿和名称缓存解析后的集合；扫描期间检查取消。`NamedListValidationTest.NamedList_ShouldCacheSourcePerWorkbook`、`NamedListValidationTest.NamedList_ShouldCacheValuesPerWorkbook` 验证同工作簿复用和跨工作簿隔离；未增加大型性能矩阵。

#### FIX-007（MEDIUM / SHOULD_FIX）：COMPLETED

新增采购类别、仓库、状态三字典与 ValueMap 编码往返，以及分页小计后的真实坐标、模板同名冲突、辅助 Sheet 顺延和 300 字符字典文本测试。NPOI XLS/XLSX、ClosedXML XLSX 的职责级测试均通过；现有非法名称和资源限制测试继续覆盖格式边界。

#### FIX-008（MEDIUM / SHOULD_FIX）：COMPLETED

MiniExcel 与 SpreadCheetah 对仅 `LookupList` 或仅 `ListSourceName` 的请求均在输出前明确拒绝。两个 Provider 各新增同步/异步流和文件入口直接测试，断言结构化错误、旧目标保留、调用方流和临时文件状态；双 TFM 各 8/8 通过。

#### FIX-009（MEDIUM / SHOULD_FIX）：COMPLETED

新增字典路径直接覆盖预取消、提交失败、旧目标保护和临时文件清理；已有异步写入取消、观察器次数和流所有权回归与本轮常规全量测试复用。更新 `symbol-test-map.md` 和本执行记录。新增文档 C# 示例使围栏数变为 20，文档测试已按代码内容识别采购与上传示例，逐段编译执行通过。

### Round 1 汇总

- MUST_FIX：2/2 完成；SHOULD_FIX：7/7 完成；PARTIAL、BLOCKED、FAILED：0。
- Release 全量构建：0 错误；仅有既存 `Microsoft.Bcl.Memory 10.0.9` 的 net6.0 兼容性警告。
- 受影响职责测试：NPOI、ClosedXML、MiniExcel、SpreadCheetah 均在 net6.0/net8.0 通过；文档测试 net8.0 19/19，通过全部 20 个 C# 示例围栏。
- `dotnet test Bing.Offices.sln -c Release --no-build --no-restore -m:1 --filter Category!=Large`：最终运行全部通过。首次运行的文档围栏计数与示例拆分失败已修复并复跑，不作为最终结果。
- 公共 API：Review 修复阶段未新增或删除公共成员；沿用本任务已审批的双 TFM +14/-0 差异与打包消费者证据。
- 严格 UTF-8/BOM/LF/末尾换行检查、`git diff --check`：通过。未提交、推送或发布。
- 下一步：重新进行独立 Review；填写区域保护、级联字典、分批导入扩展及分组分页/透视报表仍属于后续独立阶段。

### Round 2

- Review 状态：NEEDS_FIX；Fix Scope：recommended。
- 本轮完成 FIX-010、FIX-011、FIX-012；原 `review.md` 未修改。

#### FIX-010（HIGH / MUST_FIX）：COMPLETED

NPOI 显式列表解析在 `Formula1` 为空时继续读取 HSSF 的 `ExplicitListValues`，保留名称列表优先解析。新增 `InlineListImportTest.WorkbookImport_ExplicitList_ShouldAcceptAndRejectValues` 和 `InlineListImportTest.EntityImport_ExplicitList_ShouldValidateFixedAndDetailCells`，XLS/XLSX 均验证合法值、非法值和原始坐标。

#### FIX-011（MEDIUM / SHOULD_FIX）：COMPLETED

Workbook 与 Entity 导入均将调用方取消令牌传至名称字典解析与逐行扫描。`NamedListCancellationTest.ValidateCell_NamedListCanceledDuringScan_ShouldThrow` 在字典首行读取后确定性取消，XLS/XLSX 均验证下一行扫描中断；既有导入调用链与预取消回归通过。

#### FIX-012（MEDIUM / SHOULD_FIX）：COMPLETED

NPOI 与 ClosedXML 各新增成功异步字典工作簿读回、中途取消保护旧文件与清理临时文件的职责测试；NPOI 名称定义测试增加联合区域和真实 XLSX 外部引用 Unsupported 场景。定向运行 NPOI net6.0/net8.0 各 41/41，ClosedXML net6.0/net8.0 各 12/12。

### Round 2 验证与交付

- `dotnet build Bing.Offices.sln -c Release --no-restore -m:1`：通过，0 错误；仅有既存 `Microsoft.Bcl.Memory` net6.0 警告。首次并行构建因本机输出目录写入权限失败，串行提升权限后通过。
- `dotnet test Bing.Offices.sln -c Release --no-build --no-restore -m:1 --filter Category!=Large`：全部通过；NPOI 双 TFM 各 776 项，ClosedXML 双 TFM 各 239 项，Docs 19 项，ProviderContract 双 TFM 各 298 项；其他项目也全部通过。
- 八个生产包以本地临时后缀 `2.0.0-review2` 从最终源码重新打包；net6.0/net8.0 Consumer 均输出 `package-consumer-ok` 且版本为 `2.0.0-review2`。在线 NuGet 源不可达，本地包与本机已有依赖缓存完成还原；未修改源码版本号或发布。
- 双 TFM 公共 API 合同测试包含在全量通过结果中；本轮仅修改内部签名和测试，沿用已审批的 +14/-0 快照。额外 API 捕获工具长时间无输出已停止，未改写基线。
- `verify-encoding.mjs` 检查 61 个改动文本文件通过；`git diff --check` 通过。生产符号测试映射已更新。
- 执行方完成本轮 recommended 范围；是否 Review PASS 由下一轮独立审查确认。未 commit、push、发布。
