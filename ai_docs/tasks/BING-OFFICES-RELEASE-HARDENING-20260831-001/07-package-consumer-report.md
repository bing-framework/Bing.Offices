# Package Consumer Report

状态：`VERIFIED`

待验证当前本轮本地 nupkg/snupkg、精确版本 assets、XML docs、metadata 主 API、mapping、CSV、DI 和文档代码。报告必须区分本地 Bing 包源与第三方依赖的可信源/预热缓存，不把旧缓存当作当前包证据。

当前证据：三个 `2.0.0` 包已生成到 `artifacts/packages-vnext`；Abstractions、Core、Npoi nupkg 均包含 DLL、XML 和 nuspec。Docs consumer 未引用生产 `ProjectReference`，通过项目配置中的本地 Bing feed 加 `https://api.nuget.org/v3/index.json` 在隔离 `NUGET_PACKAGES` 缓存中还原、构建并运行，结果为 `11/11`；`project.assets.json` 将三个 Bing 包解析为 `type=package`、版本 `2.0.0`。测试覆盖 DI、mapping、CSV、文档 fence 以及 metadata XLS/XLSX 重开；仅有既有 `RegexAttribute` 过时警告。未修改项目 lock 文件。

Round 6 复验命令：

```powershell
$packages=Join-Path $env:TEMP 'bing-offices-docs-packages-round6e'
dotnet restore tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj --force --no-cache --ignore-failed-sources --source https://api.nuget.org/v3/index.json -p:RestorePackagesPath="$packages" -p:RestoreFallbackFolders=
dotnet build tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj --no-restore -p:RestorePackagesPath="$packages"
dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -c Release -f net8.0 --no-build --no-restore --logger "console;verbosity=minimal" -p:RestorePackagesPath="$packages"
```

Round 6 结果：restore/build/test 均成功，Docs consumer `11/11`；隔离缓存中三个 Bing 包均为 `package` 类型和 `2.0.0`，本地 nupkg 内容复核通过。公共源由项目配置的本地 feed 与命令行 nuget.org 源共同提供，空缓存仅使用本地 Bing feed 不能还原第三方依赖。
