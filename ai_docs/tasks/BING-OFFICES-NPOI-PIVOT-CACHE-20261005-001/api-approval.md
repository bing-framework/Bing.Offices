# NPOI XLSX 透视实现与零公共成员变化审批

- taskId: BING-OFFICES-NPOI-PIVOT-CACHE-20261005-001
- approvedBy: user（持续授权按 TODO 推进 F-003；复用已批准透视公共定义）
- approvedAt: 2026-10-05 (Asia/Shanghai)
- 范围：仅复用现有公共类型、构造函数、接口、签名及能力位，不新增或删除公共成员。

## 成员边界

NpoiPivotSourceValidator/SourcePlan/LayoutBudget、CacheWriter/Records/RecordWriter、TextEncoding、TableWriter、PageFieldWriter 均为内部类型。NpoiReportWriter、NpoiSheetContentWriter 和字典/级联目标避让的改动也为内部实现。

NpoiExcelExporter 声明既有 PivotTables/AdvancedPivotTables 能力，仅表示 XLSX 新建；不解释为 XLS/HSSF 透视、已有透视模板更新、公式计算或客户端数值验收通过。

## 验证与身份

成员差异预期双 TFM 均为零。先捕获新候选并与现行基线逐成员比较，再绑定真实源码、程序集和新生成的 pivot-feed-01 包；保留比较前基线及原始差异。比较成功前不更新仓库基线，不用覆盖基线掩盖变化。

九个候选包和普通/第三方四次隔离消费者运行已经通过。零成员差异、最终身份比较及集成门禁结果后续记录在 execution.md 与实际 artifacts；本文件不提前宣称未执行的检查成功。

F-002 文本 Count 客户端问题、当前 NPOI 首开/刷新/重开、F-004～007 和环境验收保持未完成。成员兼容与文件结构通过不能替代上述验收。
