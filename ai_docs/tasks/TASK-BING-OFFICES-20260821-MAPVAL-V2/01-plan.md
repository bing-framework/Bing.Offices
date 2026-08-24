# TASK-BING-OFFICES-20260821-MAPVAL-V2 阶段执行表

状态值：`TODO`、`IN_PROGRESS`、`BLOCKED`、`DONE`。

| 阶段/任务 | 状态 | 依赖 |
| --- | --- | --- |
| MAPVAL-000 基线、状态、兼容模式 | `DONE` | 无 |
| MAPVAL-001 API、调用链、格式盘点 | `DONE` | MAPVAL-000 |
| MAPVAL-100 Mapping Descriptor/Merger/Plan | `PARTIAL` | MAPVAL-001 |
| MAPVAL-101 Excel/CSV/固定动态统一接入 | `PARTIAL` | MAPVAL-100 |
| MAPVAL-200 Validation Descriptor/Attribute | `PARTIAL` | MAPVAL-101 |
| MAPVAL-201 Unique pending journal | `PARTIAL` | MAPVAL-200 |
| MAPVAL-300 双模型方向 Profile | `DONE` | MAPVAL-100 |
| MAPVAL-301 Profile Registry/DI | `DONE` | MAPVAL-300 |
| MAPVAL-400 JSON/XML normalized v2 | `PARTIAL` | MAPVAL-100、300 |
| MAPVAL-401 配置安全与错误路径 | `PARTIAL` | MAPVAL-400 |
| MAPVAL-500 Provider SPI/生产 IVT | `PARTIAL` | MAPVAL-101、200、400 |
| MAPVAL-501 Public extension/API approval | `DONE` | MAPVAL-500 |
| MAPVAL-600 文件拆分与迁移闭环 | `PARTIAL` | Phase 1-5 |
| MAPVAL-601 文档/Docs Consumer | `DONE` | MAPVAL-600 |
| MAPVAL-602 Benchmark/GC/完整验证 | `PARTIAL` | MAPVAL-601 |

不得通过修改本文件伪造完成状态；实际过程写入 `02-progress.md` 和 `execution.md`。