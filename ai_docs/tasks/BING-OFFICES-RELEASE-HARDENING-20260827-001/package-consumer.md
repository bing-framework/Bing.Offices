# 本轮包消费证据

## 包产物

本地源：`artifacts/packages-vnext`

- `Bing.Offices.Abstractions.2.0.0.nupkg` / `.snupkg`
- `Bing.Offices.Core.2.0.0.nupkg` / `.snupkg`
- `Bing.Offices.Npoi.2.0.0.nupkg` / `.snupkg`

包由当前工作树分别执行 `dotnet pack -c Release -o artifacts/packages-vnext` 生成，未执行 NuGet publish。

## 独立 consumer

`tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj` 使用三个精确 `2.0.0` `PackageReference`，并追加 `artifacts/packages-vnext` 为本地源；未使用项目引用。Restore 资产记录了三个 `2.0.0` 包路径：

- `tests/Bing.Offices.Docs.Tests/obj/project.assets.json`

验证命令：

```powershell
$env:NUGET_PACKAGES='F:\Data\NuGetPackages'
dotnet restore tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj --force-evaluate --no-cache --source .\artifacts\packages-vnext --ignore-failed-sources -p:RestoreFallbackFolders=
dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -f net8.0 -c Release --no-restore -p:RestoreFallbackFolders=
```

结果：`9/9` 通过，且未出现 `NU1601`；Docs fence、DI、Workbook Request、JSON/XML、CSV、ASP.NET Core 上传和兼容入口均从 `2.0.0` 包 API 编译/执行。

## 环境说明

机器级 NuGet 配置仍包含不存在的 `I:\Data\VisualStudio\Shared\NuGetPackages`；本轮使用会话级 `NUGET_PACKAGES` 与 restore 参数完成隔离验证，未修改用户配置。
