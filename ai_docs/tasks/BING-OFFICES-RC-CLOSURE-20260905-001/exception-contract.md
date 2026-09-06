# 异常合同

| 场景 | 当前行为 | 目标行为 | 状态 |
| --- | --- | --- | --- |
| 参数 null/range/invalid | 保留 ArgumentException 家族 | 原样保留 | VERIFIED |
| 取消 | 公共边界显式 rethrow | `OperationCanceledException` 原样 | VERIFIED |
| OOM/StackOverflow | catch filter 排除或清理后原样重抛 | 不吞、不转换 | VERIFIED |
| Excel/CSV 业务失败 | 公共边界包装现有 Offices exception 并 Observe | 同一实例、一次分类/通知 | VERIFIED |
| Observer 失败 | 写入 exception.Data，主异常保留 | 主异常不被覆盖 | 单元基线存在 |
| ExportToFile write delegate | 已按内容阶段原样传播并清理 | 原始业务异常传播 | DONE，net8 定向 9/9 |
| Create/Flush/Replace/Move | 包为 FileCommit | FileCommit + cleanup 诊断 | VERIFIED |
| 静态配置 Loader | 纯解析，不观察 | 明确纯解析；DI 为推荐观察入口 | VERIFIED |
| 用户 converter/validator/relation | 带 Inner/位置 | UserExtensionFailed + Inner/位置 | VERIFIED |
| TryAddPicture | 参数先校验，仅捕获 Argument/InvalidOperation | 仅吞合同允许的可恢复错误 | DONE，XSSF/HSSF 4/4 |
| Failure metadata | sink 接收结构化诊断；无 sink 或 sink 失败写结构化 Trace | 不静默吞掉降级 | VERIFIED |

稳定逻辑只依赖 `Code`、`Operation`、`Stage`；不得依赖本地化 Message。同一 Offices exception 已被观察后再次穿过边界时，dispatcher 必须利用实例标记避免重复通知。
