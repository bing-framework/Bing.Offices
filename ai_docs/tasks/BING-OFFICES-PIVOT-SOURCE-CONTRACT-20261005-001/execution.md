<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-PIVOT-SOURCE-CONTRACT-20261005-001
AI_EXECUTION_FINISHED_AT: 2026-10-04T20:32:34Z

# 实施执行报告

## 执行结论

SOURCE-001～003 已完成。ClosedXML 透视来源先于 Value/GetString 读取及缓存创建拒绝公式和 Excel 错误；保留原始一基坐标、纯值类型、取消、流所有权和文件提交合同。公共成员新增/删除/改签名均为零，完整 API 校验通过；原成员审批保留，基线只刷新已验证候选身份。

上一 Goal 轮为有效进展：新增 TODO 审查并取得当前常规门禁终态。本轮完成剩余 API 校验、身份更新、边界测试和交付记录。整个 Excel Goal 尚未完成，后续七组功能及商业环境验收继续保留。

## 任务信息

Task ID 为 BING-OFFICES-PIVOT-SOURCE-CONTRACT-20261005-001；2026-10-04T18:20:30.865Z 注册 plan-execution。依据已批准 plan、持续 TODO 实施授权，以及历史 review-progress-20261005.md 和 review-todo-20261005.md。原审查报告不改写；最终结论由新的独立 review.md 给出。

## 计划执行情况

| 计划 | 状态 | 证据 |
|---|---|---|
| SOURCE-001 | COMPLETED | 内部校验器接入 ApplyPivotTables；全部来源公式/错误先于求值、目标/缓存创建拒绝；实际来源坐标 |
| SOURCE-002 | COMPLETED | 双 TFM 各 66 项来源职责；Provider/ClosedXML 回归及真实文件、取消、清理证据 |
| SOURCE-003 | COMPLETED | 中文 XML、迁移、审批、追溯、包消费者、API、Release 及常规门禁 |

## 已完成事项

- 来源内表头、分组、求和及未选字段公式均拒绝；支持/未知函数及模板缓存新鲜/过期/缺失一致；来源外公式保留。
- 空/纯空白/大小写重复表头、完整空行、非数值金额、所有来源字段错误值均有原始坐标；空组和部分空金额允许，数字/日期/布尔/文本组类型保留。
- 字段 OrdinalIgnoreCase 比较且不 trim；精确带空格字段完整写出，缺失字段定位来源首个表头。展开行预算为 long。
- 后一个透视失败保全整个旧流/旧文件，临时文件清理、观察一次；预取消、生成中及实际暂存写入中取消均直测。
- 新内部类型放入 Provider Exports/PivotTables，无新生产友元、依赖、提交算法或公式计算选项。
- 完整来源/缓存断言、中文 XML、迁移与定位约定、生产符号测试映射均已补齐。

## 部分/未完成事项

本阶段核心计划无遗留。客户端首次显示/刷新、跨客户端兼容、多透视资源预算、独立页眉/签字锚点、高级透视、NPOI 缓存、已有透视模板、前向失败工作簿、低内存 Entity、Entity 失败工作簿和商业环境验收为后续范围。下一阶段 BING-OFFICES-PIVOT-CLIENT-LIFECYCLE-20261005-001/plan.md 已写入，不以计划存在宣称实现完成。

## 修改文件

生产：ClosedXmlPivotSourceValidator.cs（新增）、ClosedXmlReportWriter.cs、ExcelPivotTableDefinition.cs（remarks）。测试：PivotTableSourceContractTest.cs（新增）、双 TFM Consumer Program.cs。文档：docs/excel/pivot-tables.md、09-providers.md、nuget-migration.md。交付：本任务审批/追溯/审查/执行记录及根 API 基线身份。全局 dirty diff 还含此前阶段大量修改，不能全部归入本任务；外部 Utopa 未修改。

## API/数据/配置变化

双 TFM 八个程序集的全部成员记录、数量、签名哈希与原基线相同。原始字节保存在 artifacts/api-before/baseline-before.json，api-member-diff.json 两个 TFM 均为空。临时候选保留 approvedBy=user:pivot-table-mvp-2026-10-04、approvedAt=2026-10-04；完整比较通过后才复制根基线，根文件与已校验候选字节一致。候选审批定位本任务 api-approval。

来源公式/错误行为明确收紧：应用或显式计算接口先计算，再写字面值。金额仍由引擎双精度存储，业务负责精度及舍入；无新公共 API 或持久化格式。

## 测试结果

| 验证 | 最终结果 | 产物目录（均在 artifacts 下） |
|---|---|---|
| 来源第一次完整修复 | 双 TFM 各 50/50 | source-final-tests |
| 最新补充边界 | 双 TFM 各 66/66，0 失败/跳过 | source-followup-tests/source-followup-net6.trx、source-followup-net8.trx |
| 全套 Provider 合同 | 双 TFM 各 441/441 | contracts-tests |
| ClosedXML 回归 | 双 TFM 各 339/339 | closedxml-tests |
| Category!=Large 全门禁 | 27 TRX，4,473/4,473，0 失败/跳过；退出 0 | regular-tests |
| 身份刷新后的 API 合同 | 双 TFM 各 11/11 | api-contract-final |
| 最新文档职责 | net8 19/19 | docs-final-tests |
| 隔离包消费者 | net6 .NET 6.0.36、net8 .NET 8.0.31 均退出 0、package-consumer-ok | consumer-results.json |

常规门禁在追加 16 个边界用例前完成；随后只改测试/文档和已验证身份元数据，生产候选未变，针对新增部分重跑来源双 TFM、API 合同和 Docs。没有把常规总数改成未执行过的 4,505，也没有重复无关门禁或 Large 矩阵。

## Build/Typecheck/Lint/Format

- dotnet build Bing.Offices.sln -c Release --no-restore -m:1 -p:UseSharedCompilation=false 退出 0：0 错误、4 条既有 Microsoft.Bcl.Memory net6 提示。
- 八个生产项目依次 pack Release --no-build --no-restore -m:1 -p:UseSharedCompilation=false -p:PackageVersion=2.0.0 全成功；artifacts/packages/pack-commands.json 及各项目 pack.log 保留命令与日志。
- 消费者用 artifacts/consumer.NuGet.Config、consumer-cache/Net6 与 Net8 恢复；assets 证实隔离缓存，仅本地候选包与依赖缓存。
- ApiSnapshot capture 与完整 compare 均退出 0，实际输出 API snapshot comparison passed for net6.0 and net8.0；artifacts/api-verified/api-diff.json 双 TFM 均为空。
- artifacts/verification-summary.json 保留逐 TRX 计数、API 与消费者结果，原始日志不伪造。
- git diff --check 无空白错误，仅存量 ProfileFixtures XML CRLF 提示。六个本阶段 C# 严格 UTF-8 BOM/LF/末尾单换行和 XML 文档块解析通过；文档/JSON 为无 BOM UTF-8/LF。仅本任务生成 API JSON 规范化工具输出 BOM。

## 计划偏差

无功能偏差。新增 D-001 精确表头、错误字段及后序失败边界测试属于来源合同补齐。先对临时候选完整比较再复制基线，复制内容与已验证候选相同；源码身份范围不含测试、docs 和根基线。

## 基线问题与历史失败

1. 原审查复现双 TFM 各 3 个编译错误：SaveOptions 歧义、Path 属性遮蔽引起 Combine/GetTempPath 错误；测试缺 BOM，均已修复。不以旧 MVP 4,289 项代替当前候选。
2. 首次修复后双 TFM 各 44/48，4 失败；原 TRX 保留 artifacts/source-direct-tests。夹具 H10 未创建、精确边界断言写死 C2，修复仍保留完整来源/缓存断言。
3. 小写字段 net6 首次 48/50：缓存日期实际 ISO 文本 d 类型，原断言误用数值序列；纠正期望并保留类型/完整记录，随后双 TFM 50/50。该首次结果没有单独保存 TRX，不宣称另有原始产物。
4. 独立复核发现 Consumer 重复公式块和 using；精确去重后两个新包消费者真实编译运行成功。
5. API 脚本首次按无 BOM JSON 读取时报解码错误；改为 utf-8-sig 严格读取并只规范化本任务产物，没有修改生产 API 工具。

## 已知问题

本阶段无已确认的未修复生产问题。既有 Microsoft.Bcl.Memory net6 支持、SDK 生命周期提示保留。客户端首次可见值/刷新尚待下一阶段真实验证，OOXML 完整不证明客户端结果。

## 风险与回归关注点

来源公式迁移为本阶段实质变化，应先计算写字面值；缺失字段锚定首表头，其他来源错误实际定位；标题不 trim。多透视重复扫描/缓存待资源测量，不暗示恒定内存或自动共享。

## Reviewer 注意事项

核对真实链、完整源数据及全部缓存、后序失败目标保护、单次观察与取消矩阵；方法均在 symbol-test-map。历史 NEEDS_FIX 报告保留，最终新 review 承接修复证据；下一阶段不作为 SOURCE 完成条件，SOURCE 完成不作为整个 Goal 完成条件。

## Git 状态

保留既有未提交工作；未自动 git commit、git push、创建 PR、发布或改版本号。任务收口使用 task-finish --no-notify，不发送外部消息。
