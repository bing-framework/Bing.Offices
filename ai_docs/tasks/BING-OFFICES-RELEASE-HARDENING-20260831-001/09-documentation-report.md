# Documentation Report

状态：`IN_PROGRESS`

待同步 README、docs/excel、NuGet migration、XML docs、支持 TFM、资源边界、Stream/File 所有权、取消语义、API 迁移和发布限制。文档代码块继续由 Docs consumer 编译执行。

## 已完成

- 根 `README.md` 已改为当前实际 Excel/CSV 范围，明确 Word/PDF 不属于当前交付范围。
- `docs/excel/README.md` 已保留请求级 metadata、DOM/资源边界、原子 File API 和 Stream ownership 说明。
- `docs/excel/nuget-migration.md` 已同步实际 `AddNpoi(IServiceCollection): void`、`ExcelSetting.Default` 移除和兼容入口状态。
- 隔离 NuGet 缓存 package consumer 当前通过 `9/9`，包含 Markdown C# fence 编译执行和 README metadata 示例。
- 三程序集 Release XML docs 已存在；nupkg 均含 nuspec、DLL、XML、README 和 LICENSE。

## 未完成

API 收敛迁移示例、正式性能复现说明、第三方依赖空缓存 restore 方案和办公客户端互操作步骤仍需下一轮完成。当前构建存在 legacy API/旧 TFM 警告，文档不将其包装为零风险发布状态。
