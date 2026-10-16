# TODO 完成状态与依赖核对

核对日期：2026-10-15。业务顺序和 F/T/D/E 编号以 ADVANCED-PIVOT-TABLES/review-progress-20261005-07.md 为准，实际状态以所属任务最新 execution 与源码为准；编号不同的旧报告保持独立，不合并同号条目。下文“代码完成”不等于整个功能或发布验收完成。没有新增产品范围。

## 一、已完成的 TODO 列表

### 历史已关闭任务

Provider Contract Hardening 的 TODO.md 中下列必需阶段均已完成；final-report.md 记录后续资源、原子路径目标及校验加固闭环。记录不因本轮只改测试定位而失效。

| 原条目 | 状态 | 实现说明 |
| --- | --- | --- |
| PHASE-00 | 已完成 | 基线与重复实现核查 |
| PHASE-01 | 已完成 | 测试架构及合同 |
| PHASE-02 | 已完成 | Testing 项目骨架 |
| PHASE-03 | 已完成 | 共享模型、确定性请求与夹具 |
| PHASE-04 | 已完成 | 类型化预期快照与比较器 |
| PHASE-05 | 已完成 | Driver 与独立 Profile |
| PHASE-06 | 已完成 | Provider Contract 执行框架 |
| PHASE-07 | 已完成 | 标量、映射、Converter、ValueMap、动态列合同 |
| PHASE-08 | 已完成 | 校验、关系、资源及结构化错误合同 |
| PHASE-09 | 已完成 | 异步、取消、流、异常与原子提交合同 |
| PHASE-10 | 已完成 | 公式、样式、合并、模板与 Entity 合同 |
| PHASE-11 | 已完成 | ClosedXML Failure Workbook |
| PHASE-12 | 已完成 | ClosedXML Workbook Validation |
| PHASE-13 | 已完成 | ClosedXML Entity 动态列与资源 API |
| PHASE-14 | 已完成 | 跨 Provider 回归检查点 |
| PHASE-15 | 已完成 | Golden XLSX 治理 |
| PHASE-16 | 已完成 | API consumer、文档与迁移 |
| PHASE-17 | 已完成 | 最终验证与独立审查 |
| FINDING-REV-001 | 已完成 | 当时公共 API 与 ClosedXML 基线审批验证 |
| FINDING-REV-003 | 已完成 | Golden 治理 |
| FINDING-REV-004 | 已完成 | Entity 资源选项 |
| FINDING-REV-005 | 已完成 | 旧公共合同迁移 |
| FINDING-REV-006 | 已完成 | 富布局 Profile 合同 |
| FINDING-REV-007 | 已完成 | GoldenFixtures 编译 glob 排除 |

Capability Extension 的 TODO.md 七条保持完成：

| 原顺序 | 状态 | 实现说明 |
| --- | --- | --- |
| 1 | 已完成 | ExcelFormat 与细粒度能力描述 |
| 2 | 已完成 | 回调分批导入契约 |
| 3 | 已完成 | ExcelDataReader Provider 与 DI |
| 4 | 已完成 | XLS/XLSX/XLSB 转换、校验、资源、错误和取消测试 |
| 5 | 已完成 | 只读 Driver 与共享断言 |
| 6 | 已完成 | README、矩阵、候选、API 追溯与执行记录 |
| 7 | 已完成 | 当时测试、构建、编码和独立审查 |

### 后续路线已完成条目与子项

| 条目（保留所属任务） | 状态 | 实现说明与依据 |
| --- | --- | --- |
| F-001 | 已完成 | 独立页眉/签字区命名锚点；INDEPENDENT-PAGE-ANCHORS 的 T-001～005 及最终 review-round1-final.md 已闭环 |
| F-001/T-001 | 已完成 | 迁移示例坐标修复，文档围栏实际执行 |
| F-001/T-002 | 已完成 | 动态/计算列、原生/配置校验、名称长度与千页边界 |
| F-001/T-003 | 已完成 | 当时批准的双 TFM API 增量与候选身份 |
| F-001/T-004 | 已完成 | 普通和第三方包消费者 |
| F-001/T-005 | 已完成 | 当时常规/文档/编码与独立复核 |
| F-002 execution/T-001 | 已验证 | 数值及空值文件/缓存职责，不含客户端聚合结果 |
| F-002 execution/T-002 | 已验证 | 多轴布局、展开预算及请求复用隔离 |
| F-002 execution/T-003 | 已验证 | 资源生命周期、NPOI XLS 拒绝和 Provider 边界 |
| F-002 execution/T-004 | 已验证 | 双 TFM 定向与受影响合同回归 |
| 第七次清单 T-002a | 已完成 | runner 故障控制流 9/9；不等于真实 COM 完整验收 |
| F-006 临时预算原 TODO 第1条 | 代码/职责完成 | 输入、共享字符串正文/索引共用预算 |
| F-006 临时预算原 TODO 第2条 | 代码/职责完成 | NPOI XLS 转换 XLSX、Sheet XML 计费 |
| F-006 临时预算原 TODO 第3条 | 代码/职责完成 | 边界、空写、缩小、释放、隔离及失败清理 |
| F-006 临时预算原 TODO 第4条 | 代码/职责完成 | 读取前拒绝不受控组合；null 兼容 |
| F-006 临时预算原 TODO 第5条 | 已验证 | 双 TFM 职责、包 consumer 和候选 API，正式审批另计 |
| 本轮：Docs 仓库定位 | 已完成 | CallerFilePath 回退；仓库外程序集启动 9/9 |
| 本轮：Integration 仓库定位 | 已完成 | 相同回退；友元检查双 TFM 各 1/1，未绕过正式输入检查 |

F-004、F-005、F-007 的代码在已确定支持范围内已有实现，但整项仍未收口，列在第二节，不把功能一半冒充完成。CLIENT-001 的基础透视正常生命周期已通过，与 F-002 完整高级 Count 客户端合同分开记录。

本轮新增/修改文件单列：

- tests/Bing.Offices.Docs.Tests/SourceLayoutTests.cs（修改）
- tests/Bing.Offices.Tests.Integration/PublicApiContractTest.cs（修改）
- ai_docs/tasks/BING-OFFICES-ENTITY-TEMPORARY-BUDGET-20261015-001/execution.md（续轮记录与终态）
- ai_docs/tasks/BING-OFFICES-ENTITY-TEMPORARY-BUDGET-20261015-001/TODO.md（定位子问题更新，整项保持未完成）
- ai_docs/tasks/BING-OFFICES-ENTITY-TEMPORARY-BUDGET-20261015-001/symbol-test-map.md（定位方法追溯）
- ai_docs/tasks/BING-OFFICES-ENTITY-TEMPORARY-BUDGET-20261015-001/todo-status-20261015.md（本文件新增）

## 二、未完成的 TODO 列表

保持第七次清单的原顺序，不把 T-001、F-002 execution/T-001 等不同编号合并。

### 功能原优先级

| 条目 | 当前状态 | 上游条目及阻塞原因 | 关键依据 |
| --- | --- | --- | --- |
| F-002 P0 | 部分实现，BLOCKED_EXTERNAL | F-001 已完成；完整文本 Count/Min/Max/列轴客户端数值仍未验证通过。依赖可拥有专用实例的正常 Office 会话及根因确认，本轮按用户要求不做界面验收 | ADVANCED-PIVOT-TABLES/execution.md 的 T-005；ClosedXmlReportWriter、原始 owned-candidate-client-01.json |
| F-003 P1 | 代码已接入，交付未完成 | 依赖 F-002 高级客户端合同；正式 API 混合候选审批与统一 Release 输入仍未齐。不是“尚未实现缓存” | NPOI-PIVOT-CACHE/execution.md；NpoiPivotCacheWriter、NpoiPivotTableWriter、NpoiReportWriter |
| F-004 P1 | 受限保留/更新代码完成，正式收口未完 | 顺序在 F-003 后；F-003 收口及其他任务未审批 API 阻塞。未知依赖仍明确拒绝，不扩为无限保真 | PIVOT-TEMPLATE-UPDATE/execution.md；ExcelPivotTemplatePreserver、TemplatePivotTables |
| F-005 P1 | 限定范围实现完成，正式 API 未收口 | 顺序在 F-004 后；本任务成员已批，但整体候选混有其他任务待批成员。Provider/consumer/隔离 Release 已有证据 | EXCELDATAREADER-FAILURE-WORKBOOK/execution.md；ExcelDataReaderExcelImporter、ExcelDataReaderFailureWorkbookWriter |
| F-006 P1 | 部分实现 | 原序在 F-005 后；F-005/F-007 代码不构成其直接技术阻塞。真正总预算还依赖总额单位、DOM/实体/全局状态计费及超限终态的确定方案和成员审批；本轮不自行发明契约 | ENTITY-FORWARD-BATCH/execution.md 预算设计依赖段；ExcelEntityBatchImportOptions、ExcelTemporaryStorageBudget、两 EntityBatchImporter |
| F-007 P1 | 支持子集实现完成，正式收口未完 | 原序在 F-006 后；新增两 API 成员审批及真实客户端显示待验。整体低保留/总预算仍依赖 F-006，不将既有文件输出标为完全缺失 | ENTITY-FAILURE-WORKBOOK/execution.md；ExcelEntityImportOptions、两 EntityBatchImporter |

### 验收原顺序

| 条目 | 当前状态 | 上游及阻塞 |
| --- | --- | --- |
| T-001 | 部分取证，未整体关闭 | 基础 CLIENT-001 通过不能代替 F-002 全部高级数值/默认入口组合；依赖 F-002 客户端环境与完整逐场景核对 |
| T-002b | 未关闭 | 依赖已完成的 T-002a、T-001 的真实 Office 句柄/PID/阻塞/清理组合证据；未确认完整对应关系 |
| T-003 | 未关闭 | 依赖 F-002 核心合同和 T-001/T-002b；其他任务的独立 PASS 不能代替本任务最终 Review |

### 设计/边界原顺序

| 条目 | 当前状态 | 上游及阻塞 |
| --- | --- | --- |
| D-001 聚合边界 | 部分完成，BLOCKED_EXTERNAL | 日期/混合缓存、舍入及15位整数已直测；依赖 F-002 客户端完整数值 |
| D-002 地区与版本 | 部分完成，BLOCKED_EXTERNAL | 1900/1904 已有职责；地区/排序/版本仍需真实客户端环境 |
| D-003 总资源预算 | 未完整设计，BLOCKED_APPROVAL | 依赖 F-006 总额与终态方案；临时磁盘额度不能替代多透视/缓存/DOM/全局索引总预算 |
| D-004 分阶段取消 | 部分取证 | 来源/预检/文件取消已有职责；依赖对应 F-002～007 候选及引擎不可中断段的确定验证范围 |
| D-005 模板依赖 | 部分完成 | 依赖 F-004 收口；受限保留、缓存拆分/更新已有实现，未知扩展的额外支持须按明确清单设计 |
| D-006 批次终态 | 部分完成 | F-005/F-007 已明确不回滚、提交次序和回调异常；总预算终态依赖 F-006，业务持久化幂等由调用方负责 |
| D-007 暂存与输入容量 | 部分完成 | 行/错误/ZIP/XML/输出估算及临时磁盘限额已有实现；依赖 F-006 确定组合总状态预算 |
| D-008 能力发现 | 部分完成 | 新建与 TemplatePivotTables 已区分；完整刷新/计算/渲染声明依赖其后续实现，不能虚报 |
| D-009 真实打印 | BLOCKED_EXTERNAL | 依赖固定字体、纸张/打印环境与业务验收；本轮不做视觉校验 |
| D-010 可选扩展 | DEFERRED | 需单独批准公式重定位、绘图复制、递归级联、聚合排序/自定义聚合方案；不扩当前范围 |
| D-011 引擎持久化证据 | 部分完成 | F-003 已有 records 持久化/ZIP/重开职责；最终客户端层依赖 F-002/F-003 收口 |

### 环境原顺序

| 条目 | 当前状态 | 上游及阻塞 |
| --- | --- | --- |
| E-001 | BLOCKED_EXTERNAL | 依赖 F-002/F-003 有效候选及 WPS/LibreOffice 可用会话；Excel 证据不能替代 |
| E-002 | BLOCKED_EXTERNAL | 依赖宿主有效 Aspose 许可证、固定字体与目标部署环境；不能用无许可证拒绝冒充成功 |

### 临时预算任务未关闭边界（保留原顺序）

| 原 TODO | 状态 | 依赖及原因 |
| --- | --- | --- |
| F-006 其余资源范围 | 未完成 | 依赖 F-006 总额/终态方案 |
| 模板/Failure Workbook 与预算组合 | 未完成，当前明确拒绝 | 依赖完整额外临时文件生命周期计费方案；本轮不放开已规定拒绝 |
| 新2成员及历史候选审批 | BLOCKED_APPROVAL | 依赖所属任务逐成员批准；本轮不修改 baseline |
| PublicApiContract/Docs 门禁 | 部分完成 | 定位缺陷已修复；固定 output/release 缺 net6 的4个、net8的2个输入，且正式成员审批仍未关闭 |
| 历史外部验收 | BLOCKED_EXTERNAL | 回到所属任务规定环境，不新增 Docker/容量/客户端矩阵 |

历史 Hardening 的 FINDING-REV-002 仍为 ACCEPTED_LIMITATION（可选），后续明确支持矩阵及 Unsupported 边界；没有新增“完整 Excel 校验兼容”承诺。其余历史 closed 条目不重开。

下一轮：按优先级先恢复 F-002/T-005 的可用客户端根因验证条件；若继续仅允许代码开发，则先由维护者处理阻塞 F-003/F-005 的成员级审批和统一 Release 工件输入，再执行既有比较，不扩接口或另建基线。F-006 总预算必须先明确契约，不能越过这些前置条件临时拼出新限制。未发现本轮选定定位修复的剩余代码缺陷，不启动无效自动 Fix。

## 三、本轮完成的 TODO 列表

| 实际完成子问题 | 状态与实现 | 对应修改 |
| --- | --- | --- |
| Docs 隔离输出仓库定位 | 完成；增加编译时源码目录及 cwd 回退，真实源布局断言保持，9/9 | tests/Bing.Offices.Docs.Tests/SourceLayoutTests.cs |
| API 合同隔离输出仓库定位 | 完成；相同回退，友元检查双 TFM 2/2；缺 Release 输入仍明确失败 | tests/Bing.Offices.Tests.Integration/PublicApiContractTest.cs |

两个测试项目编译通过；没有完成新的业务功能整项，没有将缺失输入或 API 审批标为通过。仅运行上述必要定位回归，不做 UI/视觉、全解、性能或包重测。独立窄复核未发现可执行缺陷。记录与追溯改动见第一节文件清单；未清理额外配置，未新增依赖/CI/工具链/快照，未 commit/push/PR/发布。
