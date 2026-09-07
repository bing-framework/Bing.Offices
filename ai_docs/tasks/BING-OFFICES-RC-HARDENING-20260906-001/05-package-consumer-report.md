# Package Consumer 报告

## Pack

三个生产包和 symbol 包均已使用带当前项目构建的 `dotnet pack` 重新生成（避免复用旧 `obj` 中的已移除 TFM 资产）：

- `Bing.Offices.Abstractions.2.0.0.nupkg` SHA-256 `8FE0885DDECBE310C1B9363572DD0980110F468C42AE5A0739A38F1E909389FC`
- `Bing.Offices.Core.2.0.0.nupkg` SHA-256 `764489FC183B12969DD0F1D1120521AAEBD02B8C71BB8A6D085DE151AF4B6FCA`
- `Bing.Offices.Npoi.2.0.0.nupkg` SHA-256 `5568B1C874BF66FFCDA1D22A8A4D1D7A15B4D8AB49D688080F05C21563F4EBC9`

包内容已检查：Abstractions/Core 仅包含 `lib/netstandard2.0`，NPOI 仅包含 `lib/net8.0` DLL/XML、nuspec、README、LICENSE；未发现 `netcoreapp3.1`、`net6.0` 或 project 类型资产。

## Consumer 状态

`PASS`：独立 consumer 位于 `artifacts/package-consumer`，仅使用 `PackageReference`，没有 `ProjectReference`；目标 `net8.0`，使用独立缓存和 `artifacts/packages-rerun` 本地源恢复。外部 NuGet 源因当前 TLS 凭证不可用未参与本次恢复，第三方 nupkg 来自此前已验证缓存；格式化复验后运行输出：`package-consumer-ok excelBytes=4250 csvBytes=20 npoiExtensions=ok`（XLSX 字节长度受 NPOI 元数据影响，不作为 API 断言）。

`project.assets.json` 确认顶层包为 `Bing.Offices.Npoi` 与显式 `Microsoft.Extensions.DependencyInjection`，没有项目引用。格式化复验后包内 DLL 与当前 Release 输出 SHA-256 均一致：Abstractions `2FE6D84BD4D9FD8A46CF787555A1E03D851F8658412EF07ACD4C9025AA00C695`、Core `FD87EC6AC2988891723349D0550B95C603F9BC5101A724E3E37632D052439CC0`、Npoi `435000996C927F2170160D7383222F56B12149475E097336977FE21CD32B3765`。

复验命令：`dotnet pack` 三个生产项目（`-c Release --no-restore`，不使用 `--no-build`）；`dotnet restore` consumer（独立 offline source/cache，`-p:NuGetAudit=false`）；`dotnet build` 与 `dotnet run`（`--no-restore --no-build`，同一属性）。Consumer 运行时覆盖 DI、Excel/CSV、ExportMappingBuilder、JSON/XML v2 loader、`ExportToFile` FileCommit observer 同实例/单次与结构化分类，并实际访问六组 NPOI public extension 容器。最终包内 DLL 与 Release 输出 SHA-256 一致：Abstractions `2FE6D84BD4D9FD8A46CF787555A1E03D851F8658412EF07ACD4C9025AA00C695`、Core `FD87EC6AC2988891723349D0550B95C603F9BC5101A724E3E37632D052439CC0`、Npoi `435000996C927F2170160D7383222F56B12149475E097336977FE21CD32B3765`。
