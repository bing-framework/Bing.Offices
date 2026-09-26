<!-- AI_REVIEW_STATUS: PASS -->
AI_TASK_ID: BING-OFFICES-CAPABILITY-EXTENSION-20260925-001
AI_REVIEWED_AT: 2026-09-26T02:24:07.3779009+08:00

# Independent Review

## Review Scope

已对照 `plan.md`、`TODO.md`、`execution.md`、`change-impact.md`、`provider-capability-matrix.md`、`production-symbol-test-map.md`、当前 Git Diff、真实源码与最终验证产物完成独立复审。重点检查 ExcelDataReader 生产边界、完整与分批导入语义、只读 DI、XLSB 声明、资源限制、异常/取消/临时文件生命周期、公共 API、六包消费者、CI 静态门禁、编码和文档一致性。

## Statuses

```text
IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
API_STATUS: PASS
PACKAGE_CONSUMER_STATUS: PASS
CI_CONFIGURATION_STATUS: PASS
ENCODING_STATUS: PASS
EXTERNAL_GATE_STATUS: BLOCKED_EXTERNAL for Linux/container/external CI/font/production-capacity evidence
RELEASE_STATUS: LOCAL_PASS / BLOCKED_EXTERNAL
GOAL_STATUS: STOPPED_BLOCKED
OPEN_ACTIONABLE: 0
```

## Findings

### FIX-001 - NPOI XLSB/未知格式静默降级

- 处理要求：`MUST_FIX`
- 状态：`CLOSED`
- 结论：NPOI helper、同步/异步 Stream 与文件导出入口均在写入前显式拒绝 XLSB/未知枚举，不再创建 HSSF/XLS 作为降级结果。
- 证据：`ExcelHelper.ValidateFormat`、`NpoiExcelExporter.ValidateExportFormat`；`NpoiFormatBoundaryTest` 覆盖 public helper、Stream 和文件目标保全。

### FIX-002 - 缺少必需表头仍交付默认实体

- 处理要求：`MUST_FIX`
- 状态：`CLOSED`
- 结论：列绑定错误会形成错误批次并在实体读取前停止；不再交付默认值实体。`ExcelBatchImportSummary.IsSuccess` 同时考虑 `ErrorCount`。
- 证据：`ExcelDataReaderBoundaryTest` 的同步/异步缺失表头与可选表头测试；生产批次路径统一使用相同绑定结果。

### FIX-003 - DataRowStartIndex 零基/一基偏移

- 处理要求：`MUST_FIX`
- 状态：`CLOSED`
- 结论：是否进入正文统一使用零基行索引，诊断与批次首行继续使用一基物理行号。
- 证据：完整导入、同步批次和异步批次的 spacer/non-zero header 直接测试均通过。

### FIX-004 - 分批动态列与跨行 Unique 未预检

- 处理要求：`MUST_FIX`
- 状态：`CLOSED`
- 结论：`ValidateBatchPlan` 在暂存和回调前拒绝动态列及 Unique 规则，返回结构化 `UnsupportedFeature`，零批次交付。
- 证据：配置动态列、属性 Unique、同步/异步边界测试均验证 Provider/Stage、回调次数及调用方流所有权。

### FIX-005 - 公共资源限制被部分忽略

- 处理要求：`MUST_FIX`
- 状态：`CLOSED`
- 结论：完整与分批入口先执行统一 Provider 预检；`MaxSheets` 按整个 Workbook 统计。XLSX 的 `MaxCells`/`MaxColumnsPerSheet` 由共享 ZIP/XML 预检按所有 Sheet 的物理 `<c>` 和最大使用列执行，覆盖未选 Sheet 与表头前行；完整导入超限返回空根，分批仅交付结构化错误批次。XLS/XLSB 无法精确提供这两项物理预算时显式返回 `UnsupportedFeature`，未静默忽略。
- 证据：`ExcelDataReaderExcelImporter.ValidateProviderPreflight`、`ScanWorkbookResourceBudgets`、`ExcelXlsxZipPreflight.ScanWorksheetStructure`；`Import_ShouldApplySheetColumnAndCellBudgetsWithoutPartialEntities`、`Import_ShouldCountWorkbookCellsOnUnselectedSheetsAndBeforeHeader`、`ImportAsync_ShouldCountWorkbookCellsOnUnselectedSheets`、`Import_XlsbPhysicalCellLimits_ShouldBeExplicitlyUnsupported`。

### FIX-006 - XLSB 只有能力声明

- 处理要求：`SHOULD_FIX`
- 状态：`CLOSED`
- 结论：仓库加入许可与来源可追溯的真实 XLSB fixture，并验证完整导入、单 Sheet 分批导入、日期/布尔/文本值和调用方流所有权。
- 证据：`tests/Bing.Offices.Testing/Resources/Golden/xlsb/Issue635.xlsb`；`Import_ShouldReadRealXlsbFixtureAndDeliverBatches`；Provider 文档记录上游 commit、SHA-256 与 MIT 许可。

### FIX-007 - 分批入口泄漏第三方解析异常

- 处理要求：`SHOULD_FIX`
- 状态：`CLOSED`
- 结论：同步/异步分批解析异常统一包装为带 `Code/Provider/Operation/Stage/InnerException` 的公共异常；调用方回调异常保持透传边界。
- 证据：`ImportBatches_MalformedWorkbook_ShouldUseStructuredProviderException` 和 callback exception 直接测试；同步/异步共用相同结构化异常分派。

### FIX-008 - 取消与暂存清理证据不足

- 处理要求：`SHOULD_FIX`
- 状态：`CLOSED`
- 结论：暂存路径具备可测试内部 seam，所有退出路径在 `finally` 清理。测试覆盖首次真实异步读取后取消、回调取消、回调异常和损坏输入；调用方输入流保持打开。
- 证据：`ImportAsync_MidReadCancellation_ShouldKeepCallerSourceOpen`、`ImportBatchesAsync_CallbackCancellation_ShouldPropagateAndCleanStaging` 及 staging cleanup 直接测试。

### EVIDENCE-001 - 旧 API/consumer 产物曾落后于当前候选

- 处理要求：`MUST_FIX`
- 状态：`CLOSED`
- 结论：生产源码冻结后重新 build/pack 六包并生成 r5 API identity/compare；消费者在全新 source、obj、bin 和 NuGet cache 中 restore/build/run。独立复核 12 组 feed/cache nupkg SHA-256 及两 TFM 的六个输出 DLL 与包内对应 DLL，全部一致。
- 证据：`artifacts/api-pack-current-20260926-r4`、`artifacts/api-capture-current-20260926-r5`、`artifacts/api-compare-current-20260926-r5`、`artifacts/consumer-current-20260926-r5`。

## Acceptance Matrix

| Plan area | Result | Evidence |
| --- | --- | --- |
| Additive API、旧枚举值与双 TFM snapshot | `PASS` | `ExcelFormat.Xls=0`、`Xlsx=1` 保持；r5 net6/net8 diff 均为空，identity self-test 通过。 |
| 只读 Provider 与 DI | `PASS` | 仅注册 `IExcelImporter`/`IExcelBatchImporter`，未注册或伪造 exporter；真实包消费者验证 DI。 |
| XLS/XLSX/XLSB 固定列完整导入 | `PASS` | XLS/XLSX 直接测试和真实 XLSB fixture 覆盖映射、中文、日期、空值、缓存公式值及流所有权。 |
| 单 Sheet 分批导入 | `PASS` | 批次大小、精确/超限 `MaxRows`、Continue 错误行、表头/正文索引、回调串行等待与不回滚语义均有直接测试。 |
| 资源限制与失败原子性 | `PASS` | Workbook 级 `MaxRows/MaxSheets`、XLSX 物理 Cell/列预算、空根结果和错误批次均验证；XLS/XLSB 精确物理预算为显式受测边界。 |
| 异常、取消与临时文件 | `PASS` | 完整/分批结构化异常、真实异步读取中途取消、回调取消/异常、暂存清理与 caller-owned stream 均验证。 |
| 既有 Provider 的 XLSB 边界 | `PASS` | NPOI、MiniExcel、ClosedXML 均显式拒绝，NPOI 文件目标不会被替换。 |
| API/package/CI 本地门禁 | `PASS` | 六包 identity、隔离 consumer、精确 IVT allowlist 与 CI 六包/双 TFM 静态配置均通过。 |
| 文档、追溯与编码 | `PASS` | capability matrix、API approval、执行报告与生产符号映射一致；严格 UTF-8/BOM/EOL 和 `git diff --check` 通过。 |

## Verification Reviewed

- Release solution build：`artifacts/verification-final-20260926/build-final.log`，exit `0`，0 warnings / 0 errors；独立的 ApiSnapshot tool build 也为 exit `0`。
- Full solution test：`artifacts/verification-final-20260926/full-test.log`，exit `0`，合计 `2456/2456` passed，0 failed / 0 skipped。
- 各职责计数：Core `204/204`、共享 Integration `30/30`、Docs `10/10`、NPOI `575/575`、MiniExcel `44/44`、NPOI Integration `30/30`、MiniExcel Integration `9/9`、ClosedXML `102/102`、ClosedXML Integration `4/4`、ProviderContract `193/193`、ExcelDataReader `32/32`；除 Docs 仅 net8.0 外，其余均分别在 net6.0/net8.0 通过。
- API：`api-identity-r5.log` 与 `api-compare-r5.log` 均 exit `0`；`api-diff.json` 在 net6.0/net8.0 下均为空；PublicApi IVT 精确名单包含且只包含 `Bing.Offices.ExcelDataReader.Tests`。
- Package consumer：net6/net8 restore/build/run 均 exit `0` 并输出 `package-consumer-ok`；六个产品包在两套独立 cache 中与 r4 feed 原始 SHA-256 相同，consumer 输出 DLL 与 nupkg 对应 entry 逐项一致。net6 仅有 SDK 的 `NETSDK1138` 生命周期警告。
- Hygiene：最终 focused 检查 `checked=394 violations=0`；Reviewer 另抽核本任务 35 个文本文件的严格 UTF-8 解码、BOM/EOL 和末尾换行，均符合仓库规则；解决方案保留基线已有 BOM；`git diff --check` exit `0` 且无输出。
- Linux、容器、外部 CI、字体和生产规模容量未在当前 Windows 工作区伪造，保持 `BLOCKED_EXTERNAL`，不计入 `OPEN_ACTIONABLE`。

## TODO Classification

```text
Completed:
- Additive API、ExcelDataReader 只读 Provider、XLS/XLSX/XLSB 完整导入与单 Sheet 分批导入
- FIX-001 ... FIX-008
- 双 TFM API identity/snapshot、六包消费者、CI 静态门禁与生产符号追溯
- Release build、2456/2456 full test、UTF-8/EOL/diff hygiene

Open Actionable:
- none

Blocked Approval:
- none

Blocked External:
- Linux/container/external CI/font/production-capacity evidence

Not Applicable:
- none

Accepted Limitations:
- ExcelDataReader 不导出，不支持 Entity、动态列、关系、图片、Workbook 原生校验和 Failure Workbook
- 公式只读取缓存值；已交付批次不回滚
- XLS/XLSB 请求 MaxCells/MaxColumnsPerSheet 时显式 Unsupported；XLSX 使用物理 Cell 预检

Verified Boundaries:
- caller-owned input stream
- 只读 DI 不伪造 exporter
- 暂存文件在成功、失败、取消和回调异常后清理

Deferred:
- none

Next Action:
- STOP；仅在目标环境补充 Blocked External 证据
```

结论：计划范围内的本地实现、测试、API、包消费者、CI 配置和交付证据均已闭环，既往 8 个 finding 与候选身份证据问题全部关闭。当前 `OPEN_ACTIONABLE=0`，独立审查结果为 `PASS`；仅保留无法由本机替代的外部环境验证项。
