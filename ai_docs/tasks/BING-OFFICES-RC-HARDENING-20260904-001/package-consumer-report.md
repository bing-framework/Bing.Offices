# Package Consumer Report

## Consumer 结构

- 项目：`ai_docs/tasks/BING-OFFICES-RC-HARDENING-20260904-001/artifacts/package-consumer/Consumer.csproj`
- Target Framework：`net6.0;net8.0`
- 生产包引用：`Bing.Offices.Npoi` `2.0.0`、`NPOI` `2.7.4`、`Microsoft.Extensions.DependencyInjection` `8.0.0`
- 无 `ProjectReference`；`EnableDefaultCompileItems=false`，只编译 consumer `Program.cs`。
- NuGet source：任务目录 `../packages` 加 `https://api.nuget.org/v3/index.json`，配置见 `NuGet.Config`。
- `project.assets.json` 结构化检查：`projectReferences = null`，`libraries` 中无 `type=project` 条目。

## 覆盖内容

consumer 实际运行并检查：

- `AddBingOfficesNpoi` 返回原始 `IServiceCollection`；DI 可解析 Excel/CSV importer/exporter。
- Excel round trip、CSV round trip 和统一异常元数据。
- `ExcelDateAttribute.OffsetPolicy`/`OffsetMinutes`。
- Cell、Row、Sheet、Workbook、CellStyle、Font 六类 NPOI public extensions。
- HSSF/XSSF format 分支。

## 结果

| TFM | Restore | Build | Run | 输出 |
| --- | --- | --- | --- | --- |
| `net6.0` | PASS | PASS | PASS | `package-consumer-ok` |
| `net8.0` | PASS | PASS | PASS | `package-consumer-ok` |

## 最终本地包身份

| 包 | SHA-256 | NuGet content hash |
| --- | --- | --- |
| `Bing.Offices.Abstractions.2.0.0.nupkg` | `7E403F74C27288EBC8C80F2F6437F878E2F2B9B86A3C04A7DB9A08E3265CDA2B` | `WRmvNzUYZeUoFhK/GTtHvHZwM8y6IsOgNz8Jti78qsHPGogdNpu/+/J2wXts3gRZqOGknGAJ+X2wQfg5eXXG7Q==` |
| `Bing.Offices.Core.2.0.0.nupkg` | `075A42E2EE48E107DFD5CB2FDA1E8DD9EA2441364616495F5C93C25BDEB556CD` | `Fnlg2oCjBIwXEvQY3nAse8PIZjKB4X/JcUOtmHAs0aogEU9JqIMU/AKWssqgNTTv2iunrX8akUIzCstwhCUHLw==` |
| `Bing.Offices.Npoi.2.0.0.nupkg` | `C226F955C54E1838896DD087A0D883006871A910558069A3D409F52EDC1BD7CE` | `hJgvCmQmj1sN1EBLA6j2ZYSAL5BI6ocDQmmH6uf2aUvep0zrj8L3kbcusyLXW6brIMsZO0i4o9t39L7HBJG9tw==` |

NPOI nupkg 的 `netcoreapp3.1`、`net6.0`、`net8.0` asset 均已确认包含当前 `MaxReadColumns` 成员。

## 历史问题与边界

- 首次 package-only net6 运行曾因 stale multi-target output 抛 `MissingMethodException`，原因是 nupkg 的 net6/netcoreapp3.1 DLL 仍是旧 `MaxColumnLength`，而 net8 已是新成员。
- Review 修复后的首次 restore 仍复用了旧 global-packages 同版本缓存，导致 net6 consumer 找不到 `Bing.Offices.Dates`；任务 nupkg 和当前 Release 输出的 Core DLL 哈希一致，而 consumer 旧缓存 DLL 哈希不同。
- 使用任务专属全新 `.packages-final` 缓存恢复、编译和运行后，net6/net8 均输出 `package-consumer-ok`；最终 lockfile 已记录本次 nupkg 的新 content hash。
- 同版本 `2.0.0` 本地 nupkg 被重新生成，属于当前工作区交叉验证，不是发布源、clean clone 或 feed immutability 证明。
- 历史深路径 consumer 缓存曾触发 SDK/MSBuild `MSB3106`；本次任务内 consumer 使用任务相对源和隔离缓存成功，环境限制仍需在发布前处理。

## 结论

功能级 package-only consumer：`PASS（net6/net8）`；发布交付证明：`PARTIAL`，仍受正式 API approval、feed/clean-clone 和环境复现条件限制。
