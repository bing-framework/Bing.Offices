# 本轮 TODO 与未关闭边界

本文件记录 F-006 的增量，不覆盖历史清单，也不把旧检查项全部自动勾选。

- [x] NPOI/ClosedXML 前向 Entity 输入、共享字符串共用临时磁盘预算。
- [x] NPOI XLS 转换 XLSX 和逐 Sheet XML 计入同次预算。
- [x] 精确边界、同步/异步、空写、缩小、释放、调用隔离、取消与 IO/回调失败清理。
- [x] 不受控组合读取前 Unsupported；null 保持原有行为。
- [x] 两 Provider 双 TFM 职责测试、包消费者与 API 候选。
- [ ] F-006 其余资源范围：引擎缓存、DOM、完整实体结果与其他全局状态的组合内存预算。本轮临时磁盘预算不能替代。
- [ ] 模板预检与 Failure Workbook 同预算组合：当前选择性预检拒绝，需要覆盖额外临时文件生命周期后才能放开。
- [ ] 两个新增公共成员的正式审批与 baseline 更新；历史 27 条候选独立处理。
- [ ] 在可定位仓库的环境重跑 PublicApiContract/Docs SourceLayout 门禁；2026-10-15 续轮已修复定位，两方法的源布局 9 项及友元双 TFM 2 项通过；正式快照双 TFM 仍因统一 Release 输入缺失未通过，不能将整项勾选。
- [ ] 历史 Office 客户端、Docker/容量和商业 Provider 等外部验收保持原任务状态，本轮不标记完成。

续轮已完成子问题：Docs SourceLayoutTests.GetRepositoryDirectory 与 Integration PublicApiContractTest.FindRepositoryRoot 的隔离输出定位修复；无生产代码或 API 变化。完整历史顺序、完成与依赖汇总见 todo-status-20261015.md。
