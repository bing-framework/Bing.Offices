<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-BATCH-RELATIONS-20260930-001
AI_EXECUTION_FINISHED_AT: 2026-09-30T13:24:00.000Z
AI_EXECUTION_STARTED_AT: 2026-09-30T13:01:38.675Z

# 执行记录

任务：`BING-OFFICES-BATCH-RELATIONS-20260930-001`

状态：已完成。

## 变更

- ExcelDataReader Workbook 同步/异步分批路径允许既有 `HasMany(...)` 关系请求。
- 新增 Provider 内部批次关系协调器：成功实体在回调前进入内部聚合根，所有 Sheet 读取完成后统一绑定，支持父子 Sheet 乱序和 `ICollection<T>` 导航集合。
- 父行转换/校验失败或父项缺失时，子批次继续交付；关系错误以最终子 Sheet 空数据诊断批次报告，并合并到对应摘要。
- 关系错误遵守错误预算；资源/错误上限提前终止时跳过关系阶段，取消沿用同步/异步现有边界。
- 更新 ExcelDataReader Provider 能力文档、限制说明和生产符号到测试方法追溯；未新增公共 API。

## 验证

- ExcelDataReader Release 常规测试：net6.0 60/60、net8.0 60/60，筛选 `Category!=Large`。
- 关系专项测试：net6.0 3/3、net8.0 3/3，覆盖 XLS/XLSX 乱序关系、异步父行失败和缺失父项诊断。
- Docs 测试：net8.0 19/19，筛选 `Category!=Large`。
- 受影响 Provider 构建通过，0 警告、0 错误；已执行 `git diff --check`。
- 改动 C# 文件完成 UTF-8 BOM、LF 和末尾换行检查；公共 API 未增加成员。
