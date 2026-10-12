<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-REPEATED-PAGE-ANCHORS-20261003-001
AI_EXECUTION_FINISHED_AT: 2026-10-03T15:49:50Z
AI_EXECUTION_STARTED_AT: 2026-10-03T15:07:53.885Z

# 实施执行报告

## 结果

新增 `PageSubtotalNamed`，为非末页的分页小计生成连续工作簿名称（如 `PageTotal_001`），可供公式和外部程序定位。NPOI 的 XLS/XLSX 与 ClosedXML 的 XLSX 行为一致；与分组小计组合时，名称指向实际分页小计行。导入仍按原有 marker 边界工作。旧布局调用及构造签名未变。

模板已有同名锚点时，导出预检名称范围、所属工作表与历史多余名称；合法锚点随实际页位置更新，非法或过期名称在输出前报错。空明细和仅一页不产生分页名称。

## 验证

- 职责测试：NPOI net6.0/net8.0 各 19 项，ClosedXML net6.0/net8.0 各 11 项，全部通过。覆盖真实工作簿、XLS/XLSX、分组组合、模板全局及局部名称、异常、异步与导入往返。
- 受影响 Provider 全套：NPOI 各 854 项，ClosedXML 各 282 项，ProviderContract 各 298 项，Docs 19 项，全部通过。
- `Bing.Offices.sln` Release 构建通过；`Category!=Large` 全量常规测试双 TFM 通过。保留现有 Microsoft.Bcl.Memory 的 net6.0 支持警告。
- 八个生产 NuGet 包按当前版本打包。net6.0、net8.0 消费者均从包构建并运行通过，覆盖新增 API 和实体导入往返。
- API 快照比较：net6.0/net8.0 各增加 2 个成员、删除 0 个；成员审批及快照基线已更新，最终比较通过。
- `git diff --check` 和本任务改动文件的 UTF-8、BOM、行尾、末尾换行检查通过。

## 范围与后续

只处理重复页命名锚点；未改动其他待办能力。下一阶段可独立设计重复页模板/页眉与锚点的组合语义，再评估透视报表等较大能力。未提交、推送、发布或修改版本号。
