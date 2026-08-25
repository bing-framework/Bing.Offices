# NuGet Migration

当前 major 使用 `MIGRATION_CURRENT_MAJOR`。三个包身份和依赖方向保持不变：

`Bing.Offices.Abstractions <- Bing.Offices.Core <- Bing.Offices.Npoi`

迁移建议：

1. 新代码使用四种方向明确的 Profile 契约，按方向配置 Builder。
2. JSON/XML 使用 v2 `ExcelMappingDocument`；旧平铺 v1 必须通过 `MigrateV1Json`/`MigrateV1Xml` 显式选择方向。
3. 使用 `ExcelRequired`、`ExcelRegex`、`ExcelDate`、`ExcelMaxValue`、`ExcelRange`、`ExcelMaxLength`、`ExcelUnique`。
4. 继续通过 `AddNpoi(): void` 注册 NPOI；Profile Registry 使用独立的显式或程序集扫描扩展。
5. 只依赖 provider-neutral 请求、结果和转换器接口，不引用 NPOI 类型。
6. 调用方提供的输入、输出和配置流仍由调用方拥有，库不会关闭它们。

旧 Profile 快照、`object profile` 工厂重载和旧 DI 注册入口已从 next-major 公开面删除。当前任务不执行包发布；本地 pack consumer 是验收步骤。
