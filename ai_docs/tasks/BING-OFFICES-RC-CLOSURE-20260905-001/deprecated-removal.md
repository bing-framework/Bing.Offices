# 弃用删除台账

| 删除范围 | 当前引用证据 | 替代路径 | 状态 |
| --- | --- | --- | --- |
| `ExcelMapping.For<T>` | 生产定义及所有运行时/测试/文档调用已删除 | 方向化 Import/Export builder + Workbook/Sheet builder | VERIFIED |
| `ExcelMappingBuilder<T>` / old column builder | 三个旧 public 类型及 API 台账已删除 | `ImportMappingBuilder<T>` / `ExportMappingBuilder<T>` | VERIFIED |
| DataTable `Bing.Offices.CsvHelper` | 已删除生产文件、专属测试、API 台账和生成 XML 残留 | `ICsvImporter`/`ICsvExporter` + options/stream extensions | VERIFIED |
| public `MigrateV1Json/Xml` | public 重载和专属私有解析器已删除；runtime 仅接受 v2 | 升级前离线转换；runtime 仅 v2 parse | VERIFIED |
| `LegacyCompatibility` 测试/叙述 | 专属 Docs test 已删除 | 唯一 v2 推荐路径 | VERIFIED |
| v1 parsing Benchmark | 字段、setup 与两个场景已删除 | v2 parse、plan/cache baseline | VERIFIED |
| 专用于旧 API 的 DTO/helper/DI/reflection/sample/baseline | 最终反向扫描无可编译残留 | 删除或迁移到 v2 主链 | VERIFIED |

删除纪律：先确认符号与反向调用链，再迁移有效测试和文档，最后删除实现；不得误删七个 NPOI public 扩展。最终全仓 `rg` 必须无运行时、测试、Benchmark、Docs 残留，并提供 Breaking 迁移示例。

当前残留扫描只允许 `docs/excel/nuget-migration.md` 中作为 Breaking Change 键出现的 `ExcelMapping.For<T>`；不得存在可编译示例或运行时成员。
