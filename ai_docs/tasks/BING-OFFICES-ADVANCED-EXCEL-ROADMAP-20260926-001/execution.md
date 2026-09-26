<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-ADVANCED-EXCEL-ROADMAP-20260926-001
AI_EXECUTION_FINISHED_AT: 2026-09-26T21:47:15+08:00

# 开源 Provider 业务能力实施报告

## 当前范围

按用户本轮批准的 F01–F06 执行。此前报告关于五条路线全部完成的概括不作为本轮验收依据；Aspose、公式重新计算、高级 Custom 校验及透视表延期，不修改实现，不宣称通过。

## 影响分析

- ChangedProjects：Abstractions、Core、SpreadCheetah、NPOI、ClosedXML、MiniExcel、ProviderContract、Integration、Consumer、StreamingProbe。
- ChangedPublicContracts：图片与原生数据校验类型、Sheet 描述和 Builder 增量成员。
- ChangedRuntimePaths：映射、绘图/校验写入、报表预检、暂存、IO。
- ChangedTFMs：netstandard2.0 公共层，net6.0/net8.0 Provider。
- ChangedBuildPackaging：Provider 合同新增 SpreadCheetah 引用，Docker 固定环境和容量 Probe。
- RiskLevel：HIGH；先职责定向测试，再受影响项目和最终全量门禁。

## 当前进度

| 项目 | 状态 |
| --- | --- |
| F01 映射与背压 | COMPLETED；SpreadCheetah 57/TFM，含 17 映射用例、边界/背压/故障/取消 |
| F02 图片与原生校验 | COMPLETED；SheetContent 83/TFM，含 PNG/JPEG、32 比较、列表、日期系统、拒绝与清理 |
| F03/F04 报表 | COMPLETED；NPOI Report 25/TFM、ClosedXML Report 11/TFM；条件规则、预检、汇总行和作用域成功读回 |
| F05 合同/consumer/API | COMPLETED；全量 2820/2820、双 TFM API compare、八包 fresh-cache 消费者全部通过 |
| F06 Docker/容量 | COMPLETED；最终容器 352/352、fontless 真实对照、9/9 最终容量通过 |

## 实现偏差

- SpreadCheetah 1.28.0 仅原生嵌入 PNG，drawing 像素换算与公共 96 DPI 不一致。仅有图片请求暂存文件，修正小型 drawing 及 JPEG 媒体关系；JPEG 原始字节封装，不解码重编码。无图片路径直接前向输出。临时文件 DeleteOnClose，目标流归调用方。
- NPOI 2.7.4 将显式列表引号写为 CDATA 内实体字面量。仅新增显式列表的 XLSX 暂存后流式修正 dataValidations 节点，不载入完整 worksheet，不升级依赖。
- 通用 Stream 不承诺回滚；路径复用 IFileExportCommitter。图片暂存不替代原子提交算法。
- 既有 net6.0 Microsoft.Bcl.Memory 10.0.9 兼容性警告保留，不掩盖、不升级依赖。

## 状态分类

- Completed：F01–F06、最终全量构建测试、API/包消费者、容器与容量；4 项 Review 修复独立关闭。
- Open Actionable：无。
- Blocked Approval：正式性能阈值未批准；本轮只实测。
- Blocked External：远程 CI 未执行；不影响本轮已要求的本地 Docker 验收，不触发重测。
- Not Applicable：新图片流式 API 的同构历史 before 性能比较。
- Accepted Limitations：Stream 非原子、Provider 前向/DOM 边界。
- Verified Boundaries：无字体 NPOI AutoFit 结构化失败、SpreadCheetah 成功；旧文件原子保全、Stream 非原子失败边界。
- Deferred：Aspose、重新计算、高级校验、透视表。

## Git 与通知

保留原有未提交改动。未 commit、push、创建 PR、发布或升级版本；不发送外部通知。

## Review 修复记录 — Round 1

Fix Scope=recommended，独立报告 `review.md`；补充复核后共 3 MUST_FIX + 1 SHOULD_FIX。

- FIX-F03-001：NPOI HSSF 普通条件格式颜色分支由独立实施子任务修复，补无色/前景/背景/双色、CellValue/Formula、同步/异步保存读回。
- FIX-F01-001：补固定/动态 Getter、Converter、格式/类型转换的结构化错误和位置元数据；保留取消、已结构化异常和致命异常，数据源枚举/目标 IO 异常不误包为转换失败。
- FIX-F04-001：SpreadCheetah 链接公共报表预检，所有 Sheet 的 Table/WorksheetOptions 在创建输出前准备；重复名称、越界、样式、筛选冲突均提前失败。16 场景/TFM 已通过。
- FIX-F04-002：SpreadCheetah 同范围 Table/Filter 合并已读回验证；NPOI 将筛选语义保留于 Table，避免额外 Sheet 筛选，由子任务修复。
- 同轮直接补充：有界名称/打印区域正向检查、重复标题边界、冻结起点越界、默认色阶红/绿与蓝数据条。36+4 个 net8.0 直接用例通过。
- 初次全量测试运行期间主代码与新增回归测试继续变化，出现旧 Provider DLL 与新测试的混合结果；该轮 `full-tests` 是阶段诊断证据，不作为 Final Candidate。冻结后重跑唯一最终全量门禁。
- F06-D01–D04 工具缺口已修：真实 fontless 对照、CI smoke/证据上传、保留 HOME、容量独立临时目录。`review-docker.md` 是修复说明，最终独立复核由未实施这些改动的 Reviewer 完成。
- 容量证据失效范围调整：新增表格实际行数检查及映射异常包装改变了行循环，旧 9 样本只保留为历史。冻结后重新运行一次 9 单元最终容量，而不是因为阈值未批准反复测量。

- F04 最后两项直接核查：NPOI ShowTotals 原先未追加汇总行，现保留全部业务行并追加空行，表格范围扩大一行、筛选范围不变；共享预检拒绝末行溢出。名称写入先设置 SheetIndex，再赋名称，修复全局/局部同名被错误视为全局重复的问题。保存读回确认作用域 -1/0/1；NPOI Report 25/TFM、ClosedXML Report 11/TFM 通过。
- Final Candidate 已冻结。`final-build.log`：Release 构建退出 0、0 错误；`final-test.log` 与 25 个最终 TRX：2820/2820、0 失败/跳过，未过滤 Large。Docker smoke 352/352；9 个容量单元全部成功，最高工作集 81,268,736 bytes。

## 最终验证与交付证据

- L0/L1/L2：职责定向、受影响双 TFM 完整项目测试通过，方法映射见 `production-symbol-test-map.md`。
- L3：Provider 合同 276/TFM；最终包消费详见 `package-consumer-report.md`。
- L4：全量 Release build/test 通过，详见 `integration-report.md`。
- L5：`docker-final/streaming-capacity.jsonl`，9/9 EXPECTED_SUCCESS，单进程串行、4 GiB；阈值不由本轮自行批准，详见 `resource-report.md`。
- 格式：161 个相关文本严格 UTF-8/BOM/EOL 检查无问题、`git diff --check` 通过；证据 `final-encoding.json`，保留无关存量格式契约。
- 历史证据：初次混合源码全量日志、旧容器合同和旧容量不用于最终 PASS。报告更新后复用冻结代码结果，不重复运行未受影响门禁。
- API/包收口：全量测试期间并行打包发生制品竞争，测试结束且输出稳定后重新 no-build 打包八包；只刷新批准 snapshot 的 candidateIdentity，不改变 API 形状，不覆盖历史 before。最终 capture/compare 位于 `api-final/capture`、`api-final/compare`，双 TFM 空 diff；fresh-cache consumer net6.0/net8.0 restore/build/run 全部退出 0。

## Round 1 收口

Implementation=6/6；本轮本地门禁=8/8（职责/受影响项目、构建、全量测试、API、consumer、Docker、容量实测、编码差异）。独立 Review=PASS，OPEN_ACTIONABLE=0。No-Progress Check=CHANGED；Next Action=STOP。本轮计划状态 COMPLETED；产品正式发布仍保留远程 CI/性能阈值审批边界，不启动额外修复或重复容量循环。
