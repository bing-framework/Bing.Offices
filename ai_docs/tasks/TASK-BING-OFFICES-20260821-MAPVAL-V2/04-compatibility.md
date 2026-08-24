# TASK-BING-OFFICES-20260821-MAPVAL-V2 兼容矩阵

## Task 信息

- Task ID：`TASK-BING-OFFICES-20260821-MAPVAL-V2`
- Compatibility Mode：`MIGRATION_CURRENT_MAJOR`

## NuGet 包矩阵

| 包 | 当前身份 | 依赖方向 | 迁移策略 |
| --- | --- | --- | --- |
| `Bing.Offices.Abstractions` | 保留 | 无 Core/Npoi 依赖 | additive public contracts；生产 IVT 清理后仅暴露 SPI/消费者契约 |
| `Bing.Offices.Core` | 保留 | `Abstractions` | 旧 mapping/validation API 委托新 compiler；internal 化实现 |
| `Bing.Offices.Npoi` | 保留 | `Core` | `AddNpoi(): void` 保持；通过 SPI 消费 Core Plan，不依赖生产 IVT |

## 旧 API 矩阵

| 旧 API | 当前策略 | 删除时机 |
| --- | --- | --- |
| `ExcelMappingProfile<T>` | 保留并标记 Obsolete，委托新双方向模型 | 下一 major，须有迁移证据 |
| `ExcelMappingConfiguration` | 保留读取/兼容 facade，内部归一化到 v2 document | 下一 major 或独立生命周期确认后 |
| `IExcelMappingConfigurationLoader` | 保留名称和现有重载，新增演进重载 | 不删除已发布签名 |
| `AddNpoi(): void` | 保持签名，分步注册 Profile | 明确 next-major 前不改返回类型 |
| 旧 Attribute | 保留并 Obsolete，统一 Descriptor Factory | 下一 major |
| `ExcelStreamExtensions` File/Bytes | 保留 | 不删除 |
| `CsvStreamExtensions` 同步 File/Bytes | 保留 | 不删除 |

## JSON/XML 矩阵

| 格式 | 读取 | 输出 | 诊断 |
| --- | --- | --- | --- |
| v1 平铺 `columns` / 缺省 version | 必须继续 | 不再新输出 | migration diagnostic，默认不阻断 |
| v2 `version/profile/import/export` | 必须支持 | 新输出只写 v2 | 精确 JSON/XML 路径 |
| XML DTD/外部实体 | 拒绝 | 不产生 | `XmlResolver = null`、DTD Prohibit |
| 未知字段/超限/非法 validator | 按版本策略拒绝或诊断 | 不输出 | 不加载任意 CLR 类型 |

## 当前状态

- 包 build/pack consumer：已通过本地验证；locked restore 仍因历史 lock 文件与项目依赖范围不一致报告 `NU1004`，执行使用 `--force-evaluate`，未改写受跟踪 lock 文件。
- Public API approval：已由 `PublicApiContractTest` 验证 allowlist、成员快照哈希和生产 IVT 约束。
- v1/v2 真实互操作：JSON/XML v1/v2 normalized loader、XML DTD 拒绝、未知字段和输入限额已有测试；round-trip writer、诊断 API 和 model alias registry 尚未完成。