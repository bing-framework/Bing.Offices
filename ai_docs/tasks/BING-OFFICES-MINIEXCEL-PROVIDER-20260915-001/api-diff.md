# API 差异与包治理

## Additive surface

- 新增程序集/包：`Bing.Offices.MiniExcel`。
- 新增公共类型：`MiniExcelExcelExporter`、`MiniExcelExcelImporter`、`ExcelMiniExcelServiceCollectionExtensions`。
- Abstractions/Core/NPOI 只做 XML 注释中性化、共享预检内部抽取和既有测试回归；未删除公共成员。
- MiniExcel public surface 不暴露 `MiniExcelLibs` 或 NPOI 类型。

## 验证

```text
dotnet run --project .\build\ApiSnapshot\ApiSnapshot.csproj -c Release --no-build -- --capture true --root .\output\release --output .\artifacts\api-snapshot-final-pack --packages .\artifacts\packages --repository .
dotnet run --project .\build\ApiSnapshot\ApiSnapshot.csproj -c Release --no-build -- --root .\output\release --packages .\artifacts\packages --repository . --output .\artifacts\api-verify-final
```

结果：`API snapshot capture completed for net6.0 and net8.0`；`API snapshot comparison passed for net6.0 and net8.0`。

包验证结果：4 个 `.nupkg`、4 个 `.snupkg`；MiniExcel 包包含 net6.0/net8.0 DLL/XML、README、LICENSE，依赖 Core、DI Abstractions 和 `MiniExcel [1.46.0]`，无 NPOI dependency/tag。

`build/api-snapshot-baseline.json` 的 `approvedBy=task-plan-additive-api` 用于记录本任务的 additive baseline；它不是人工成员审批凭证，发布前仍应由仓库维护者确认。
