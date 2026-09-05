<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BING-OFFICES-RC-HARDENING-20260904-001
AI_REVIEWED_AT: 2026-09-05T20:37:24.9429159+08:00

# 独立验收 Review

## 验收摘要

最终结论：`NEEDS_FIX`。

本次为 Round 3 Review Fix 后的复审。基于当前源码调用链、Git Diff 和直接回归测试，`FIX-001`、`FIX-002`、`FIX-003` 均继续保持解决状态，未发现与这些修复直接相关的新回归。Round 3 没有开放的 `MUST_FIX`，未产生业务或测试代码变更。`FIX-004`、`FIX-005` 仍未纳入 `fixScope=must`，当前源码仍未满足对应计划合同，均保持 `NOT_RESOLVED`。

当前没有未解决的 `MUST_FIX`，但仍有 2 项 `SHOULD_FIX`，因此按 Review 协议不能判定 `PASS_WITH_ISSUES`。发布结论仍为 `BLOCKED / No-Go`；formal API approval、性能/资源预算和完整发布证据也尚未解除。

Reviewer 仅更新本文件，未修改业务代码、测试代码、`plan.md` 或 `execution.md`。

## Review 边界

- 计划范围：Abstractions/Core/NPOI 的异常合同、确定性日期、XLSX 预检、NPOI Provider User API、Breaking API 治理、测试、包消费、Benchmark、资源探针和发布门禁。
- 当前工作区为共享 dirty worktree：46 个已跟踪文件产生约 2541 additions / 456 deletions，另有任务目录、异常/日期实现、ZIP preflight 和测试等未跟踪文件；没有独立 commit 边界。Round 3 仅更新 `execution.md`。
- Diff 与本任务主体一致，但无法仅凭 Git 证明每处修改的起点；本次按计划、源码调用链、测试和 execution 历史共同归属。
- `git diff --check` 未报告 whitespace error，仅报告 CRLF/LF 转换提示。
- `get_errors` 对 `src` 和 `tests` 未发现编辑器诊断错误。

## 上一轮 FIX 复核

### FIX-001 复核

- 状态：`RESOLVED`
- 原问题：`MaxZipCompressionRatio = null` 时错误访问 `Nullable.Value`。
- 当前证据：`ExcelResourceLimits.Validate()` 只在 `HasValue` 时读取压缩比，并拒绝非正、NaN 和 Infinity。
- 直接测试：null 关闭和全部无效值均包含在三 TFM 专项中。
- 回归判断：未发现回归。

### FIX-002 复核

- 状态：`RESOLVED`
- 原问题：嵌套 catch、反射和 NPOI 包装层会把取消或致命异常转换、吞掉或降级为结构化数据错误。
- 当前证据：CSV/NPOI converter、validator、setter、relation delegate 和 Failure Workbook 路径显式传播 `OperationCanceledException`，并排除 `OutOfMemoryException`、`StackOverflowException`。
- 包装处理：`CsvPropertyBinding`、`ExcelColumnPlan`、`NpoiRelationBinder` 使用 `ExceptionDispatchInfo` 解包 `TargetInvocationException`；`NpoiFailureWorkbookWriter` 通过内部异常链恢复 serializer 包装的取消或致命异常。
- 测试覆盖：Excel/CSV converter、validator、setter、relation delegate、serializer 和 diagnostic sink，并断言原异常实例与取消令牌。
- 实际验证：本轮 net8/net6/netcoreapp3.1 相关专项各 `55/55` 通过。
- 回归判断：普通用户扩展异常和可恢复数据错误的既有断言同时通过；未发现回归。

### FIX-003 复核

- 状态：`RESOLVED`
- 原问题：XLSX DOM 前预检缺少独立 XML 字符数和嵌套深度预算。
- 当前证据：`ExcelResourceLimits` 提供默认 `MaxXmlCharacters=64 MiB`、`MaxXmlDepth=256`，支持 null 关闭并拒绝非正配置。
- 生产实现：`NpoiXlsxZipPreflight.ValidateXmlSafety()` 使用禁用 DTD/entity 的流式 reader，累计节点名、节点值、属性名和属性值，检查 `reader.Depth`，并以 `BingOfficesResourceLimitException`、`Stage=Preflight` 报告超限。
- 主链接入：`NpoiExcelImporter.ImportCore()` 在 `WorkbookFactory.Create()` 前调用生产 preflight。
- 测试覆盖：workbook、sharedStrings、styles、worksheet 四类 XML 的字符/深度超限、深度边界、null 关闭和无效配置。
- 实际验证：本轮三 TFM 相关专项各 `55/55` 通过；net8 独立进程 ResourceProbe `1/1` 通过，验证 `xml-depth-limit`、`xml-character-limit` 在 `Preflight` 拒绝且 `importedRows=0`。
- 回归判断：未发现回归。

### FIX-004 复核

- 状态：`NOT_RESOLVED`
- 原问题：`SheetExtensions.TryAddPicture()` 两个 public overload 的参数和异常合同不完整。
- 当前证据：`IPictureData` overload 仍直接解引用参数；byte[] overload 仍只校验 `sheet`，没有显式拒绝空字节、负 row/col 或无效 `PictureType`，并继续使用 broad `catch (Exception)` 返回 false。
- 判断：上一轮明确跳过，当前源码未达到修复目标，也未发现等价替代实现。

### FIX-005 复核

- 状态：`NOT_RESOLVED`
- 原问题：日期导入缺少多输入格式和 request/document/profile 独立配置面。
- 当前证据：全仓生产配置仍无 `InputFormats` 或 `DateParsing`；`ExcelDateAttribute` 只有单个 `Format`；`ExcelDateParser` 对 `DateTime` 只调用单格式 `TryParseExact`，请求级 Sheet/Workbook API 没有日期输入格式配置。
- 判断：上一轮明确跳过，当前源码未达到计划要求。

## 计划逐项验收矩阵

| Phase / Task | 结果 | 实际证据 |
| --- | --- | --- |
| RC0-01 至 RC0-03 基线与矩阵 | PASS | 基线、进度、API/弃用报告存在；当前 Git 状态和 Diff 已复核 |
| RC1-01 异常合同 | PASS | `FIX-002` 已由生产调用链和三 TFM 直接测试闭环 |
| RC1-02 日期合同 | PARTIAL | 确定性 ISO、offset、1900/1904 主链存在；多格式与请求级独立配置缺失，见 `FIX-005` |
| RC1-03 API Contract | PARTIAL | 重复键已消除且测试可运行；formal API hash 仍待维护者批准 |
| RC1-04 XLSX ZIP 预检 | PASS | DOM 前 ZIP/XML 预算、四类 XML 矩阵和独立进程 probe 已验证 |
| RC1-05 Failure Workbook / Validation | PARTIAL | 异常合同和主要行为已验证；完整双 DOM 资源矩阵仍未批准/完成 |
| RC2-01 NPOI public extensions | PARTIAL | 六类扩展已 public 且 consumer 可见；`TryAddPicture` 公共边界不完整，见 `FIX-004` |
| RC2-02 / RC2-03 Mapping 与 rename | PASS | 死字段和目标 rename 已落地；formal approval 仍独立阻断 |
| RC2-04 兼容面清理 | PARTIAL | 已批准子集完成；剩余兼容策略仍待治理 |
| RC3-01 至 RC3-05 职责拆分 | PARTIAL | 计划明确的完整拆分、目录整理和 A/B 未完成，执行报告已记录偏离 |
| RC4-01 至 RC4-04 验证与 consumer | PARTIAL | 修复专项、Build、历史 Integration/Docs/consumer 有证据；formal API hash 仍失败 |
| RC5-01 至 RC5-04 Benchmark / Resource | PARTIAL | MappingValidation 和现有 probe 有证据；完整矩阵与预算审批未完成 |
| RC6-01 文档 | PARTIAL | 迁移与执行文档存在；剩余公开合同修复后仍需同步 |
| RC6-02 API approval | NOT_VERIFIABLE | 需要维护者审批，Reviewer 不更新 formal baseline |
| RC6-03 独立 Review | FAIL | 仍有 2 项 `SHOULD_FIX` 未解决 |
| RC6-04 发布门禁 | FAIL | 公开合同、API approval、预算和完整发布矩阵未全部解除 |

## Git 与实现 Review

- `NpoiXlsxZipPreflight.Validate()` 在 `WorkbookFactory.Create()` 前真实接入，不是空接口或仅测试实现。
- XML 预算与 ZIP entry/部件字节预算职责独立；ResourceProbe 的限制模式调用真实 importer。
- 异常修复覆盖反射包装和第三方包装，不依赖异常消息分类。
- 六类 NPOI extension 已形成真实公共 API 扩张；正式 API baseline 未被自动修改。
- 未发现 reset、clean、commit、push、tag 或 publish 操作证据。

## API、架构与资源 Review

- `BingOfficesException` 层次和稳定元数据已形成公共合同；参数、取消、致命异常与可恢复数据错误的主要边界现已一致。
- Core 日期 parser 是 Excel/CSV 共用的单一解析实现，但配置模型仍不足以表达计划要求的多输入格式。
- `TryAddPicture` 当前 broad catch 仍可能吞掉未知或致命异常，且参数错误与可恢复 NPOI failure 无法区分。
- formal API candidate 与正式 hash 不一致仍是有效发布门禁，不得通过更新测试期望绕过审批。
- Failure Workbook 双 DOM、100k/1M、取消延迟、LOH retained 和完整 Benchmark/ResourceProbe 预算尚未形成可批准发布证据。

## 测试与文档 Review

- 前次复审完整 Release solution build：PASS，0 errors，16 warnings；Round 3 无业务或测试代码变更，未重复运行完整构建。
- 本轮 FIX-001/002/003 专项：net8 `55/55`、net6 `55/55`、netcoreapp3.1 `55/55`。
- 本轮独立进程 ResourceProbe：net8 `1/1`。
- execution 记录的完整 Unit：每个 TFM `453/454`，唯一失败为 formal API snapshot hash；Integration net8/net6 各 `15/15`。
- 本轮未重复运行完整 Unit、Integration、Docs、package consumer 和 Benchmark；相关结论采用 execution 证据并与当前源码、专项测试和前次完整构建交叉核对。
- `execution.md` 已记录 Round 2 修复范围和剩余阻断；`FIX-004`、`FIX-005` 修复后仍需同步对应文档和 consumer 证据。

## 问题分级

- BLOCKER：0。
- HIGH：0。
- MEDIUM：2，见 `FIX-004`、`FIX-005`。
- LOW：0；不新增纯风格问题。

## FIX-004

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 当前状态：`OPEN`
- 对应计划项：`RC2-01`、`RC4-04`。
- 涉及文件/符号：`SheetExtensions.TryAddPicture()` 两个 public overload 及 package consumer/tests。
- 问题：公开扩展只校验 `sheet`，未校验 `pictureData`、`pictureBytes`、负 row/col 或无效 `PictureType`；实现捕获所有异常并返回 false。
- 证据：`IPictureData` overload 直接读取 `pictureData.Data`；byte[] overload 仍存在 `catch (Exception)`。
- 影响：公共参数错误的异常类型不稳定，调用方无法区分无效输入与可恢复 NPOI failure；取消或致命异常可能被静默吞掉。
- 修复目标：标准参数错误立即抛出；Try 模式只对文档声明的可恢复失败返回 false；取消和致命异常原样传播。
- 明确修复要求：校验两个 overload 的所有参数，缩小 catch 范围，更新 XML docs 的 returns/exceptions，不得以维持 false 结果为由吞掉未知异常。
- 修复后的验证方式：增加 null、空字节、负索引、无效枚举、正常 HSSF/XSSF、可恢复失败和取消/致命异常传播 Unit；package-only consumer 覆盖正常和参数错误路径。

## FIX-005

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 当前状态：`OPEN`
- 对应计划项：`RC1-02`、`RC4-01`、`RC6-01`。
- 涉及文件/符号：日期配置 DTO/attribute/fluent/mapping/options、`ExcelDateParser`、Excel/CSV conversion/validation、Docs/consumer。
- 问题：当前仅支持 `ExcelDateAttribute.Format` 单格式；Mapping 动态规则也只有单个 `Format`，请求级列配置没有独立日期输入设置。
- 证据：生产源码没有 `InputFormats`/`DateParsing`；parser 只读取单个 `attribute.Format`；请求级 Workbook/Sheet API 未提供日期格式集合。
- 影响：调用方无法在不修改实体 attribute 的情况下，按 request/document/profile 显式接受多个合法日期形态；未满足批准计划的确定性日期配置合同。
- 修复目标：保持默认仅 ISO `yyyy-MM-dd`，同时允许调用方在不修改实体 attribute 的情况下显式配置一个或多个输入格式；导入格式与导出 Formatter/NumberFormat 分离。
- 明确修复要求：沿用现有 Mapping merge/snapshot 模型；支持格式顺序、校验、克隆、合并、JSON/XML 和 package consumer；不得退回宽松 `DateTime.Parse` 或本机 Culture/时区行为。
- 修复后的验证方式：Excel/CSV 覆盖 `yyyy/MM/dd`、`yyyyMMdd`、中文日期的默认拒绝和显式多格式接受；验证 validation/conversion 一致、request 覆盖 document/profile、snapshot 隔离、DateTimeOffset policy 和 1900/1904 不回归；运行三 TFM Unit、Integration、Docs 和 package consumer。

## 未完成与外部阻断

- formal API candidate 尚未获得维护者批准，正式 baseline 不得修改。
- 性能/资源预算未批准，完整 Benchmark 与 ResourceProbe 矩阵未完成。
- 剩余兼容 API 和 public execution detail 的治理策略未完全收口。
- 本地同版本 nupkg consumer 成功不等于正式 feed、clean clone 或 publish 证明。

## 回归与兼容风险

- 收紧 `TryAddPicture` 参数和异常行为属于公共 behavior change，必须同步测试、文档和 package consumer。
- 日期多输入格式会改变 public API candidate，必须重新生成成员级 API diff 并等待维护者批准。
- XML 字符/深度默认预算是新的公共配置面；正式发布前仍需在批准的资源样本上验证默认值兼容性。
- 所有代码问题完成后仍不能自动批准 formal API baseline、资源预算或发布状态。

## 最终验收 Checklist

- [x] plan、execution、旧 review、当前 Git Diff 和真实源码已读取。
- [x] `FIX-001` 已独立验证为 `RESOLVED`。
- [x] `FIX-002` 已独立验证为 `RESOLVED`。
- [x] `FIX-003` 已独立验证为 `RESOLVED`。
- [x] 三 TFM 关键专项和 net8 独立进程 ResourceProbe 已通过。
- [x] 完整 Release solution build 已通过。
- [x] `git diff --check` 无 whitespace error，编辑器无 src/tests 诊断错误。
- [ ] NPOI public extension 参数和失败合同完整。
- [ ] 日期多输入格式和请求级独立配置满足计划。
- [ ] formal API approval、完整 Benchmark/resource 和发布环境证据完成。
- [ ] 独立 Review 的 `SHOULD_FIX` 清零。

## 最终裁决

`FIX-001`、`FIX-002`、`FIX-003` 已解决；`FIX-004`、`FIX-005` 仍为 `SHOULD_FIX / OPEN`。因此最终状态保持 `NEEDS_FIX`，后续可交由 `review-fixer` 按现有 FIX ID 继续处理，Reviewer 不自动修复。
