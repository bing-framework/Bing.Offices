# Resource Report

- Task ID：`BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001`
- 原始矩阵：`artifacts/resource-probe/resource-matrix.jsonl`
- 可读报告：`artifacts/resource-probe/resource-matrix.md`

## 结果

本机 Release 运行 3 种 staging 策略（Memory、TempFile、Hybrid）、3 种场景、4 档请求并发，共 36 个 child 场景，全部 `status=passed`；实际 DOM 并行度固定为 1，排队和退出事件均被记录。最大观测值约为 allocated `3,571,693,024 B`、peak working set `119,377,920 B`、temporary disk peak `423,347 B`，所有场景 temporary after 为 0、leftover files 为 0。

MiniExcel 自身的输入复制使用真实 `ReadAsync`/`WriteAsync`，取消和 source ownership 由 MiniExcel 专项测试直接覆盖；资源矩阵验证的是共享 staging/原子提交边界及 NPOI DOM 场景。

## 审批与限制

矩阵执行返回 `scenarios=36 status=passed approval=BLOCKED`。未提供批准人和批准时间，因此不把本机资源数据宣称为已批准的 2 CPU/4 GiB 发布预算。该矩阵实现沿用了既有 ResourceProbe 的固定 `BO-RC-20260908-002` 内部 task 标识，本任务将其作为可追溯的复用证据记录；这不是当前 taskId 的新审批。

