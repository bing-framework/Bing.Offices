# 决策记录

## D-001 缺失分析与方法论输入

- 状态：已记录
- 事实：指定的 `ai_docs/codebase-analysis/bing-offices-implementation-review-20260904.md`、Mustang methodology manifest/router 当前不存在。
- 决策：以当前源码、项目配置、可复现命令和测试/基准产物为主证据；不伪造缺失报告或路由结果。
- 影响：执行报告需明确该输入缺失。

## D-002 共享工作树保护

- 状态：已记录
- 事实：当前源码和测试没有新的已跟踪改动，仅新 task 目录未跟踪；前序任务的 dirty 状态属于共享工作区。
- 决策：不执行 reset/clean/restore，不覆盖前序任务文件；本任务只增量修改计划范围内文件。

## D-003 DateOnly 范围

- 状态：计划决策
- 决策：当前 `netstandard2.0`/`netcoreapp3.1` 兼容矩阵不公开 `DateOnly`，改为明确“不支持”；除非后续获得 TFM/API 变更批准，不增加伪兼容。

## D-004 DataTable CSV

- 状态：待验证
- 决策：Phase 0 先统计真实调用和外部价值；默认删除 Core DataTable 双轨，不创建无明确消费者证据的兼容包。

## D-005 NPOI 扩展可见性

- 状态：已确认方向
- 决策：Cell/Row/Sheet/Workbook/CellStyle/Font 扩展属于 Provider-specific User API，恢复 public；`InternalExtensions`、`PictureTypeResolver` 和执行 helper 保持 internal/private。

## D-006 API baseline

- 状态：待批准
- 决策：本任务所有异常新增、NPOI 扩展公开、rename/delete 均先生成成员级 candidate diff；未获批准不更新正式 baseline，不用修改 hash 掩盖差异。

## D-007 异常翻译边界

- 状态：计划决策
- 决策：参数异常、取消和致命异常保留原语义；可恢复行级问题保留 Import Error；NPOI/CsvHelper/IO/配置公共运行失败在单一公共边界翻译为 `BingOfficesException` 子类，已包装异常不重复包装。
