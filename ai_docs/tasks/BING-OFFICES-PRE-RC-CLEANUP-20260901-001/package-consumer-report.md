# 独立 NuGet 包消费者报告

## 消费方式

最终消费者位于 `artifacts/package-consumer-rerun2`，目标框架为 `net8.0`，仅通过 `PackageReference` 引用：

- `Bing.Offices.Npoi` 2.0.0；
- `Microsoft.Extensions.DependencyInjection` 8.0.0；
- `Microsoft.Extensions.DependencyInjection.Abstractions` 8.0.0。

`Consumer.csproj` 使用 `EnableDefaultCompileItems=false` 显式编译 `Program.cs`，不包含任何 `ProjectReference`。`obj/project.assets.json` 中 `Bing.Offices.Abstractions/Core/Npoi` 均以 `package` 类型解析，且 `projectReferences` 为空。消费者覆盖 Excel、CSV、Mapping loader、DI 链式注册和 `ExcelImportFailureOptions.MaxSerializedBytes` 的构造/校验。

## 本地包身份

本次消费者使用任务目录 `artifacts/packages` 中重新生成的本地 nupkg，关键 SHA-256 如下：

| 包 | 版本 | SHA-256 |
| --- | --- | --- |
| `Bing.Offices.Abstractions` | 2.0.0 | `3EB4606C52E0C28764575E3BC6141CBC31F45DAD03BADD2BBD7020808A6A5438` |
| `Bing.Offices.Core` | 2.0.0 | `97DF1A949C08C8E38AC5AC6E2720EB4C1B801770104C5FA8392A6C2FBFADCB29` |
| `Bing.Offices.Npoi` | 2.0.0 | `0B5CC9B3D5D96F66A4F2A63AB3BBEF85D306EBDC5EDB2A6FBE323ECFF4432885` |

包内容检查确认 NPOI 程序集提供：

```text
AddBingOfficesNpoi(IServiceCollection): IServiceCollection
```

## 短路径缓存验证

为规避 SDK/MSBuild 在任务深路径下的程序集引用路径处理问题，最终成功验证使用短路径 NuGet global packages cache `C:\nupkg-cache`。`NuGet.Config` 清空默认源，并保留任务本地包目录、nuget.org 和 Visual Studio Offline Packages。

```powershell
$task = (Resolve-Path .\ai_docs\tasks\BING-OFFICES-PRE-RC-CLEANUP-20260901-001).Path
$env:NUGET_PACKAGES = 'C:\nupkg-cache'
dotnet restore "$task\artifacts\package-consumer-rerun2\Consumer.csproj" --force-evaluate --no-cache --configfile "$task\artifacts\package-consumer-rerun2\NuGet.Config"
dotnet build "$task\artifacts\package-consumer-rerun2\Consumer.csproj" -c Release --no-restore
dotnet run --project "$task\artifacts\package-consumer-rerun2\Consumer.csproj" -c Release --no-build
```

上述 restore、build、run 均退出码 `0`，运行输出为：

```text
package-consumer-ok
```

这证明当前本地 nupkg 在 `net8.0`、无 `ProjectReference` 条件下可被第三方消费者编译和运行；不扩展为其它 TFM 或发布环境的结论。

## 长路径缓存限制

在任务目录下的专属长路径缓存执行 `package-consumer-rerun`/`package-consumer-rerun2` 时，restore 成功但 build 失败，出现：

```text
MSB3106: 程序集强名称 "...Microsoft.Extensions.DependencyInjection.Abstractions.dll" 的路径找不到，或者是格式不正确的完整程序集名称
```

随后出现 `CS1069`、`CS0012` 和 `GetRequiredService` 缺失。诊断确认 `project.assets.json`、NuGet generated props/targets 和 DLL 实体均存在；切换到 `C:\nupkg-cache` 后相同 package-only consumer restore/build/run 全部成功。因此该失败记录为 SDK 10.0.300/MSBuild 长路径环境限制，不归因于 nupkg API 或依赖缺失。长路径失败后的旧二进制输出不作为证据，也不运行或采信旧 `artifacts/package-consumer` consumer 的历史输出。

## 限制与状态

- 仅验证 `net8.0` package-only consumer；缺失的 netcoreapp3.1、net5.0、net7.0 runtime 未执行对应验证。
- 包版本为工作树的 2.0.0，未执行发布；包身份仍依赖当前任务 artifact，不能替代正式不可变 feed/版本治理。
- 短路径验证为 `VERIFIED`；长路径缓存为 `BLOCKED`；consumer 总体状态为 `PARTIAL`，不解除 API baseline、性能预算和缺失 runtime 的发布阻断。

## Round 5 最新包验证

Round 5 使用 `artifacts/packages-round5` 重新生成的包，并将 `package-consumer-rerun2/NuGet.Config` 的本地源切换到该目录。消费者仍为 net8.0、仅 `PackageReference`、无 `ProjectReference`；`Program.cs` 额外编译使用 `ExcelRequiredAttribute`，以验证旧 validation attributes 删除后的替代 API 可消费。

| 包 | 版本 | Round 5 SHA-256 |
| --- | --- | --- |
| `Bing.Offices.Abstractions` | `2.0.0` | `9911681786D724482F6E19DD08F9526463CC25759DD49E399247D99AF7FB03D0` |
| `Bing.Offices.Core` | `2.0.0` | `541D39F08C8F18A3CFFB69B0F6EF64CA375CA891702AE0F927A53D6272B2A2A2` |
| `Bing.Offices.Npoi` | `2.0.0` | `67932EEB37001A3B5099D2A52FF481FC2C78B8D28A6B8AD6FFD1E2C646C42483` |

Round 5 在 `C:\nupkg-cache-round5` 中重新执行 restore、build、run，三步退出码均为 `0`，输出 `package-consumer-ok`。该结果绑定上述 Round 5 包 hash；不改变长路径 SDK/MSBuild `MSB3106` 环境限制、仅 net8.0 覆盖范围或 RC `No-Go`。
