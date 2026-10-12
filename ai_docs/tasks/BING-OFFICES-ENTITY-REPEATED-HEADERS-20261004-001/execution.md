<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-ENTITY-REPEATED-HEADERS-20261004-001
AI_EXECUTION_FINISHED_AT: 2026-10-04T12:17:38Z

# 实施执行报告

## 执行结论

Entity Layout 已支持重复打印标题行与实际插入重复页眉，NPOI XLS/XLSX 和 ClosedXML XLSX 均可导出；实体导入能校验并跳过复制页眉。Round 1 的 7 项和 Round 2 审查提出的 5 项可执行修复均已实施并通过常规门禁；公共 API 相比任务原候选只增加一个带原始坐标的配置异常构造重载，未删除或改签名。独立 Reviewer 尚需重新确认结论。

## 任务信息

- 任务：`BING-OFFICES-ENTITY-REPEATED-HEADERS-20261004-001`
- 开始：`2026-10-03T16:11:23.282Z`
- 原实施完成：`2026-10-03T17:43:50Z`
- Review 修复完成：`2026-10-04T09:28:08Z`

## 计划执行情况

P1 公共契约、P2 双 Provider/导入、P3 测试与交付均完成。`PrintTitleRows` 只写打印元数据；`PageHeader` 要求 `PageBreak`，实际复制首页连续区域、内容、公式文本、样式、行高及完整合并，按页覆写 Fluent 单元格并放置导入 marker。相同 Sheet 的两种模式互斥。

## 已完成事项

- 公共布局增加打印标题、重复页眉构建器、页码/总页数/当前页明细上下文及只读 Provider SPI 描述。
- 共享物理行规划在分页小计之后、下页明细之前插入页眉；导入扫描验证 marker 并跳过复制行。
- 两个 Provider 完成格式边界、模板占用、合并和标记预检；NPOI 的 XLS/HSSF 与 XLSX/XSSF 均覆盖。
- 文档、迁移示例、Provider 矩阵、成员审批及生产符号到测试方法映射已更新。
- net6.0、net8.0 包消费者从本地 NuGet 包执行打印标题、实际页眉与导入往返。

## 部分/未完成事项

原实施报告曾记录“无计划内未完成项”；随后独立审查发现 7 项可执行问题，已在本轮处理。公式计算、绘图对象复制和实际物理打印分页属于本计划明确排除的范围。

## 修改文件

主要涉及 `src/Bing.Offices.Abstractions/Bing/Offices/Entities`、`src/provider-shared`、两个 Provider 的 Entity 执行器、两个 Provider 测试、双 TFM 包消费者、公共 API 分类测试、`docs/excel`、`build/api-snapshot-baseline.json` 和本任务目录。未触碰外部 Utopa 类库。

## API/数据/配置变化

双 TFM API 快照各仅在 `Bing.Offices.Abstractions` 新增 26 个成员、删除 0 个；本轮增加结构化配置异常构造重载，逐成员审批见 `api-approval.md`。没有持久化数据格式或版本号变化。构造函数及原有 Fluent 调用保持兼容。

## 测试结果

- 职责测试：NPOI 重复页眉 net6.0/net8.0 各 19 项，ClosedXML 各 11 项；打印标题各 Provider 的双 TFM 测试通过。
- Release 常规全量测试使用 `Category!=Large`；除 NPOI net6.0 的既有异步写入中取消测试一次偶发失败外，其余项目均通过。该失败用例独立复跑通过，NPOI net6.0 全套 874 项串行复跑通过；NPOI net8.0 全套 874 项、ClosedXML 双 TFM 各 294 项、Provider 合同双 TFM 各 298 项通过。
- Docs 测试 19 项通过；net6.0、net8.0 本地 NuGet 包消费者均输出 `package-consumer-ok`。
- API 快照对最终源码清单、本地包和审批文件的双 TFM 比较通过；公共 API 治理测试双 TFM 各 11 项通过。

## Build/Typecheck/Lint/Format

`Bing.Offices.sln` Release 构建 0 错误；仅有存量 `Microsoft.Bcl.Memory` net6.0 支持提示。`git diff --check` 通过。本任务涉及的 29 个文本文件通过严格 UTF-8、BOM、LF 和末尾换行检查（C# 有 BOM，其他文件无 BOM）。

## 计划偏差

无功能偏差。为避免历史候选包使 API 身份校验误判，最终八个包放入本任务独立目录。消费者恢复使用本地 NuGet 源及已有缓存，未访问网络。

## 基线问题

工作区在本任务前已有大量其他任务的未提交改动，均予以保留。全量测试并发运行时，既有 NPOI net6.0 取消测试出现一次偶发失败；独立及全项目串行复跑均通过。并行编译 NPOI、ClosedXML 测试时曾遇共享 Core 输出文件锁，改为串行后通过。

## 已知问题

本轮已修复 Review 所列的已知功能故障；仍需下一轮独立 Review 判定。逻辑页仍可能与真实打印分页不同；公式只复制文本，不保证重新计算，绘图对象不复制。

## 风险与回归关注点

模板源区域必须紧邻首个明细行，marker 不可位于合并区域内部；后续页目标区域若已有模板内容或合并会在输出前报配置错误。用户应根据明细和页眉总高度预留工作表行数。

## Reviewer 注意事项

检查 `api-approval.md` 的 26 个新增成员、`symbol-test-map.md` 的 Provider 双 TFM 用例和 `artifacts/api/review-fix-candidate`、`artifacts/api/review-fix-compare` 中的候选/比较结果。全局 Git 差异包含本任务开始前的其他功能变更，不能将其全部归于本任务。

## Git 状态

未自动 git commit；未自动 git push；未自动创建 PR，也未发布或修改版本号。

## Review 修复记录

### Round 1

- Review 状态：`NEEDS_FIX`，Fix Scope：`recommended`。
- Review 文件：`ai_docs/tasks/BING-OFFICES-ENTITY-REPEATED-HEADERS-20261004-001/review.md`，原始审查结论保持不变。
- FIX-001（MUST_FIX）：`COMPLETED`。布局构建阶段拒绝绝对 marker 与固定字段重叠；两个 Provider 在命名锚点解析后也核对 marker。NPOI/ClosedXML 的固定字段冲突测试通过。
- FIX-002（MUST_FIX）：`COMPLETED`。两个 Provider 根据最终列计划检查源区域是否覆盖完整列表表头，并禁止 marker 占用表头。NPOI XLS/XLSX 与 ClosedXML 的窄源区域测试通过。
- FIX-003（SHOULD_FIX）：`COMPLETED`。两个 Provider 在模板加载后检查现有打印标题与实际页眉互斥；模板冲突测试通过。
- FIX-004（SHOULD_FIX）：`COMPLETED`。共享导入扫描器核对首页 marker、首页区域内的重复 marker 及缺失的后续页眉。NPOI XLS/XLSX 和 ClosedXML 导入负向测试通过。
- FIX-005（SHOULD_FIX）：`COMPLETED`。配置异常增加保留旧签名的坐标构造重载，scanner 使用原始 Sheet 和一基行列报错。marker 错误字段断言与双 TFM API 比较通过。
- FIX-006（SHOULD_FIX）：`COMPLETED`。构建器只限制同 Sheet 的实际页眉区域数量，允许不重叠的普通列表；NPOI 运行时足迹纳入首页源区域。两个 Provider 的并列列表真实往返通过，文档同步更新。
- FIX-007（SHOULD_FIX）：`COMPLETED`。补齐重复 marker、窄源、模板打印标题与合并冲突、同表列表、值工厂失败、NPOI 暂存文件清理及错误坐标测试；更新能力说明、成员追溯、API 审批和快照。

本轮验证：页眉职责测试 NPOI net6.0/net8.0 各 30 项、ClosedXML 各 20 项；常规 Provider 测试 NPOI 双 TFM 各 889 项、ClosedXML 在最后两个测试加入前各 303 项、合同各 298 项；Docs 19 项；双 TFM API 治理各 11 项；Release 解决方案构建 0 错误；`Category!=Large` 解决方案全量测试通过。最后只增加 ClosedXML 两项职责测试，双 TFM 定向复跑各 20 项通过，未重复全量测试。net6.0/net8.0 NuGet 包消费者均输出 `package-consumer-ok`。API Snapshot 候选与本轮包身份、审批文件比较通过，双 TFM 仅新增上述一个构造成员。

全解决方案 `dotnet pack` 因无关的 Analyzers 项目缺少 `_._` 文件报 `NU5019`，但所需八个 Office NuGet 包均生成并完成消费者运行及候选身份比较；该命令本身不能记为通过。保留既有 net6.0 `Microsoft.Bcl.Memory` 支持警告。`git diff --check` 通过，本轮改动的文本文件通过严格 UTF-8、BOM、LF 与末尾换行检查。未运行无关的大规模性能矩阵。

Round 1 汇总：MUST_FIX 2/2、SHOULD_FIX 5/5；PARTIAL 0，BLOCKED 0，FAILED 0。下一步为独立 Review。未 commit、push、创建 PR 或发布包。

### Round 2：并列列表边界与打包收口

- 补充无 `End` 的实际页眉列表与更长的相邻普通列表往返测试；修复两个 Provider 导入器按整张工作表末行扫描的问题。现在使用页眉列表自身列及 marker 列确定扫描末行，显式 `End` 和 Footer 路径保持原约定。
- NPOI 重复页眉职责测试双 TFM 各 32 项、ClosedXML 双 TFM 各 21 项通过；`Category!=Large` Provider 全套 NPOI 双 TFM 各 891 项、ClosedXML 双 TFM 各 306 项通过。
- 补齐 Analyzers 项目已声明但缺失的零字节 `_._` 包占位文件。使用本地缓存离线还原后，串行 `dotnet pack Bing.Offices.sln -c Release --no-restore -m:1` 通过；包内存在 `lib/netstandard2.0/_._`。并行打包曾因共享输出目录文件复制冲突失败，串行重跑通过。
- 使用本轮重新生成的包运行 net6.0、net8.0 消费者，均输出 `package-consumer-ok`。双 TFM 八个 Office 程序集 API 成员差异为零；更新候选源码清单及 NPOI、ClosedXML 包身份后，API snapshot 比较通过。
- 串行执行 `dotnet test Bing.Offices.sln -c Release --no-restore -m:1 --filter Category!=Large`，所有测试项目通过；保留既有 net6.0 `Microsoft.Bcl.Memory` 支持提示。`git diff --check` 与本轮改动文件的严格 UTF-8、BOM、行尾及末尾换行检查通过。
- 原 `review.md` 是修复前独立审查记录，未覆盖写。新增边界已修复并验证，仍待独立 Reviewer 复核最终实现。未 commit、push 或发布。

### Round 3：重复页眉边界与资源验收

- FIX-008（MUST_FIX）：`COMPLETED`。NPOI、ClosedXML 对每个后续页眉目标同时检查固定字段声明、已有内容及合并区域；无模板、固定值为空时也拒绝覆盖。NPOI XLS/XLSX 与 ClosedXML XLSX 职责测试核对原输出流保持不变。
- FIX-009（MUST_FIX）：`COMPLETED`。共享导入校验器在命名起点解析后重新检查页眉源末行与实际明细起点相邻；合法 `Header(false)` 命名列表可往返，起点移位返回结构化配置错误。
- FIX-010（SHOULD_FIX）：`COMPLETED`。两个 Provider 对实际页眉列表在显式 `End` 内查找最后的实际内容行，不再将预留空行算作明细；损坏 marker 仍会被拒绝。
- FIX-011（SHOULD_FIX）：`COMPLETED`。共享扫描器按页容量及分组小计占用的物理行计算缺失 marker 的位置；无分页小计、分组边界及非首页首行 marker 用例核对原始坐标。
- FIX-012（SHOULD_FIX）：`COMPLETED`。新增命名起点、无表头、动态列与计算列宽度、源高度、异步成功往返、文件提交与写入失败的直接用例。NPOI 实体文件提交异常与普通 Workbook 文件提交统一单次观察；NPOI 真实原子提交失败和 ClosedXML 默认暂存写入失败均核对原目标内容及临时文件清理。

本轮职责测试 NPOI net6.0/net8.0 各 52 项、ClosedXML 各 35 项通过。`Category!=Large` 解决方案测试全部通过，其中 NPOI Provider 双 TFM 各 911 项、ClosedXML 各 320 项、Provider 合同各 298 项、Docs 19 项通过。`dotnet build Bing.Offices.sln -c Release --no-restore -m:1` 为 0 错误；`dotnet pack` 串行成功。独立 NuGet 缓存中的 net6.0/net8.0 消费者均输出 `package-consumer-ok`。双 TFM API 成员差异为零；仅更新候选源码清单和两个 Provider 包身份后，API snapshot 比较通过。保留既有 `Microsoft.Bcl.Memory` net6.0 警告。

本轮消费者使用 `artifacts/consumer-round3.NuGet.Config` 映射独立本地包目录及本机 NuGet 依赖缓存，未访问远程源。`review.md` 保留独立审查原文；DESIGN-001～006 仍为后续独立设计，不属于本轮 fixScope。未 commit、push、创建 PR 或发布。下一步为独立 Review。

### 2026-10-07 后续：HSSF Forward 实际重复页眉

- NPOI HSSF/XLS 前向实体批次通过既有 HSSF 到受限 XLSX 事件投影复用独立实际页眉扫描；未改公共 API 或生产主链。
- 新增 `EntityForwardPageHeaderTest.ImportEntityBatches_Forward_Hssf_ShouldSkipRepeatedHeaders` 和 `EntityForwardPageHeaderTest.ImportEntityBatches_Forward_Hssf_ShouldRejectMissingRepeatedHeaderMarker`；NPOI net8.0、net6.0 各 7/7 通过，覆盖同步、异步、原始批次行号、来源行号及 marker 坐标错误。
- HSSF 输入仍只保留投影后的单元格值、行号和必要数字格式；复杂模板对象、合并、关系、原生校验以及页眉与小计或 Footer 组合继续按既有边界拒绝。
- 文档已同步说明 NPOI HSSF/XLS 的受限支持。未 commit、push 或发布。
