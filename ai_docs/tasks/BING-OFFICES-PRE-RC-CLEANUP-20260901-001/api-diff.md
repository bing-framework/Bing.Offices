# API 差异与治理记录

## 结论

### 正式 baseline 已批准

用户在本执行回合明确批准将当前 Release API 纳入正式 baseline。批准范围仅包括此前已记录并完成迁移验证的 API 收敛：

- Abstractions：`ExcelImportFailureOptions.MaxBytes` 更名为 `MaxSerializedBytes`，并删除已批准的 `ExcelSetting` / `SheetSetting`；
- Core：Round 5 已批准删除集、Round 6 已批准的 Office 异常层级和 Settings/plan/type-map/CSV concrete execution detail internal 化；
- NPOI：`AddNpoi` 删除，由链式 `AddBingOfficesNpoi(IServiceCollection)` 替代；
- Round 8：`IExcelMappingPlanFactory` Provider 的尾部可选 `cacheCapacity` 参数，用于真实缓存容量回归，不新增独立用户入口。

当前 Release 快照产物：`artifacts/api-snapshot-formal-20260903/api-snapshot-*.json`。
正式 baseline 文件：`artifacts/api-snapshot-formal-baseline-20260903.json`。

当前三个逻辑程序集的成员数/hash（netcoreapp3.1、net6.0、net8.0 一致）：

| 程序集 | 成员数 | 正式 hash |
| --- | ---: | --- |
| `Bing.Offices.Abstractions` | 723 | `7F9A2AA819E94B3838097DF2FF374A934CF7F35F3D2E91F3D1DB790F22972943` |
| `Bing.Offices.Core` | 194 | `B3661970BBE5AECC06DAD57B1E3F960FA77E70C4D2E66B2DA4910F7823AA2BB6` |
| `Bing.Offices.Npoi` | 1 | `DA163263804A964D8AC2A13D78D6B3858256171CE7729841690FDB56F602CEEE` |

与 Round 5 快照相比，逐成员差异仅为已批准删除/收敛：Abstractions 删除 `ExcelSetting` / `SheetSetting` 共 15 个成员；Core 删除 71 个旧 public execution detail 成员并新增 Provider 的 `cacheCapacity` 尾参数成员；NPOI 无新增成员差异。未发现未批准的公共 API 新增。

- Round 5 快照：`artifacts/api-snapshot-review-fix-round5/api-snapshot-*.json`

以下内容为正式 baseline 获批前的历史取证，保留用于证明当时未机械修改测试 hash；当前正式结果以本报告顶部的批准 baseline、`api-snapshot-formal-20260903` 快照和 compare 结果为准。

Round 5 在用户授权清理已迁移的 `[Obsolete]` API 后重新生成快照。该快照记录本轮真实删除；在当时的历史执行中，正式 API hash 保持旧值，Unit API 门禁按预期失败并作为 No-Go 证据保留。

## 快照比较

- Round 1 快照：`artifacts/api-snapshot-review-fix-round1/api-snapshot-*.json`
- 生成输入：`output/release` 下的当前 Release DLL；Round 1 重新生成，不使用旧 JSON 作为实际输入
- TFM：`netcoreapp3.1`、`net6.0`、`net7.0`、`net8.0`
- 工具：`build/ApiSnapshot/ApiSnapshot.csproj`

| 程序集 | TFM | 旧成员数 | 新成员数 | 旧 hash | 新 hash | 逐成员差异 |
| --- | --- | ---: | ---: | --- | --- | ---: |
| Bing.Offices.Abstractions | 全部四个 TFM | 737 | 737 | `5A1B668E14A3A2689A0CC88BB95F3396025BABC1AF18A6A5A64C0BD0DF290646` | `407B4F3C2605333A082766B13E1F1DEB704880DDE6D3E0CEDB72FF3F6281ADF0` | 删除 1，新增 1 |
| Bing.Offices.Core | 全部四个 TFM | 273 | 273 | `5F68499B76921FA52D293BC851B94659FBB2AB466E6D91C8878A97F8728B4BB7` | 相同 | 0 |
| Bing.Offices.Npoi | 全部四个 TFM | 1 | 1 | `A0DBE9808D82547601429D8958C7ED283467031A3763EB9037B19D03F19D80BD` | `DA163263804A964D8AC2A13D78D6B3858256171CE7729841690FDB56F602CEEE` | 删除 1，新增 1 |

## Round 1 产物身份

API 快照的三个逻辑程序集在四个 TFM 上相同；NPOI 程序集按目标 TFM 生成，成员快照和 hash 相同。当前 Release DLL 的文件身份如下，作为本次快照输入的可复核记录：

| 文件 | SHA-256 | 字节数 | LastWriteTime UTC |
| --- | --- | ---: | --- |
| `output/release/netstandard2.0/Bing.Offices.Abstractions.dll` | `0A07E90CB46EC9AF6D6ABB6F2E5903B412D4D415C0F2347E31103777C64A8605` | 103424 | `2026-09-02T01:31:26.5268660Z` |
| `output/release/netstandard2.0/Bing.Offices.Core.dll` | `664F5290030B5307DC9965877D6833FE50BF3B0A6FF253181FEDBDAD9FDBB207` | 141824 | `2026-09-02T02:11:23.5777355Z` |
| `output/release/netcoreapp3.1/Bing.Offices.Npoi.dll` | `CC02505667938FD2E0B72C103CB908CC4723E10E7F60FA81675A048A2CD38EBF` | 148992 | `2026-09-02T02:17:41.4641269Z` |
| `output/release/net6.0/Bing.Offices.Npoi.dll` | `5619D276A2695B9F25F1B9D54448CC3458FC3DEC3A7A1FCCAC7519EF4F5DD43C` | 143872 | `2026-09-02T02:17:41.3391245Z` |
| `output/release/net7.0/Bing.Offices.Npoi.dll` | `CF3FD811E3A910B0B949677CA5A2360CED50ACD6A8674438CE6A951EAF2C06AF` | 143872 | `2026-09-02T02:17:41.4781262Z` |
| `output/release/net8.0/Bing.Offices.Npoi.dll` | `DA4A22D5A33C8FDC5BC1CE4B0B5C5333E1B5625E71A2C3CFC9FD987B1BE40E7C` | 143360 | `2026-09-02T02:11:25.3443208Z` |

## NPOI 成员差异

```text
- method|Bing.Offices.Npoi.Extensions.ExcelNpoiServiceCollectionExtensions.AddNpoi|static|System.Void|Microsoft.Extensions.DependencyInjection.IServiceCollection|generic=0
+ method|Bing.Offices.Npoi.Extensions.ExcelNpoiServiceCollectionExtensions.AddBingOfficesNpoi|static|Microsoft.Extensions.DependencyInjection.IServiceCollection|Microsoft.Extensions.DependencyInjection.IServiceCollection|generic=0
```

## 当前工作树新增成员级差异

基于 `artifacts/api-snapshot-review10/api-snapshot-net8.0.json` 与 Round 1 重新生成的 `artifacts/api-snapshot-review-fix-round1/api-snapshot-net8.0.json` 做完整集合差异，当前结果为：

```text
[Bing.Offices.Abstractions]
- property|Bing.Offices.Imports.ExcelImportFailureOptions.MaxBytes|System.Nullable`1[[System.Int64]]
+ property|Bing.Offices.Imports.ExcelImportFailureOptions.MaxSerializedBytes|System.Nullable`1[[System.Int64]]

[Bing.Offices.Core]
	no member-level difference

[Bing.Offices.Npoi]
- method|Bing.Offices.Npoi.Extensions.ExcelNpoiServiceCollectionExtensions.AddNpoi|static|System.Void|Microsoft.Extensions.DependencyInjection.IServiceCollection|generic=0
+ method|Bing.Offices.Npoi.Extensions.ExcelNpoiServiceCollectionExtensions.AddBingOfficesNpoi|static|Microsoft.Extensions.DependencyInjection.IServiceCollection|Microsoft.Extensions.DependencyInjection.IServiceCollection|generic=0
```

`MaxSerializedBytes` 是有意的 public property rename，原因是实现只限制 Failure Workbook 的序列化输出，不限制输入解压、NPOI DOM、业务实体或输出目标流的整体内存；已完成生产代码、测试和公开文档迁移，专项测试为 `14/14`。`AddNpoi` 是有意删除并由链式 `AddBingOfficesNpoi` 替代，已完成全仓调用迁移、exact member、Integration、Docs 和 isolated nupkg consumer 验证。

本次差异生成命令使用 `artifacts/api-snapshot-rerun-baseline.json` 作为临时比较基线，退出码为 `1` 的原因是四个 TFM 的正式期望 hash 均与当前实际 hash 不匹配；该失败是 baseline 获批前的历史证据，不能通过更新临时或正式 hash 隐藏 Breaking Change。当前正式 compare 已使用 `artifacts/api-snapshot-formal-baseline-20260903.json` 并退出码为 `0`。

`5A1B...` 并不是 Round 1 当前 Release 的实际 Abstractions hash：它是 `artifacts/api-snapshot-review10` 中的旧快照值。Round 1 实际快照和 net6/net8 Unit 运行时均得到 `407B4F3C...`；成员差异显示原因是 `ExcelImportFailureOptions.MaxBytes` 更名为 `MaxSerializedBytes`。Core 的 `5F68499B...` 与旧 review10 快照一致，NPOI 的 `DA163263...` 对应 `AddNpoi` 到 `AddBingOfficesNpoi` 的成员替换。

## 取证命令

```powershell
 dotnet run --project .\build\ApiSnapshot\ApiSnapshot.csproj -c Release --no-build -- --root .\output\release --baseline .\ai_docs\tasks\BING-OFFICES-PRE-RC-CLEANUP-20260901-001\artifacts\api-snapshot-rerun-baseline.json --output .\ai_docs\tasks\BING-OFFICES-PRE-RC-CLEANUP-20260901-001\artifacts\api-snapshot-review-fix-round1
```

退出码为 `1`，原因是临时比较 baseline 仍包含旧 hash。仓库不存在 `build/api-snapshot-baseline.json`，因此没有把临时 baseline 当作正式批准基线。

## 治理决定

- `AddNpoi`：按计划 P2-03 迁移并删除旧入口；不恢复 wrapper、forwarder 或新的 `[Obsolete]`。
- `AddBingOfficesNpoi`：保留为唯一推荐 NPOI DI 入口，返回传入的同一 `IServiceCollection`。
- Abstractions/Core（Round 1 历史状态）：无成员差异，不需要修改 hash。
- NPOI hash（历史状态）：曾待维护者确认 Breaking Change 后更新；当前已按用户批准更新正式快照和对应测试。
- Round 5 已按授权删除 `ICellValueConverter`、6 个旧 validation attributes、CSV 全局 separator/quote 和旧隐式 DataTable 重载；旧异常、Settings、UniqueTracker、public execution-detail 和其它未授权 compatibility 候选继续保持未关闭，不伪造为已删除。

## 相关证据

- 本地包消费成功：`artifacts/package-consumer-rerun2`，`project.assets.json` 显示业务包均为 package 类型，运行输出为 `package-consumer-ok`。
- NPOI exact member 单元测试已改为要求 `AddBingOfficesNpoi`，并在 net8 Unit 回归中通过。
- 测试中的全程序集旧 hash 断言仍失败，详见 `unit-test-report.md`。

## Round 5：已授权 Obsolete/兼容 API 清理快照

### 产物与比较边界

- 生成命令仍使用 `build/ApiSnapshot/ApiSnapshot.csproj`，输入为当前 `output/release`；输出目录为 `artifacts/api-snapshot-review-fix-round5`。
- 覆盖 `netcoreapp3.1`、`net6.0`、`net7.0`、`net8.0`；四个 TFM 的逻辑程序集成员快照一致。
- 本轮历史 API compare 退出码为 `1`，唯一原因是临时 baseline 仍保存旧 hash；随后已在用户批准后更新正式 baseline，并以退出码 `0` 重跑 compare。

### 成员级差异

`Bing.Offices.Abstractions`：旧成员 `858`，新成员 `856`；新 hash：
`C176D71B0025C1F28F010BF05667588898A4D0EA4F847CD65658D9737D800313`。

```text
- method|Bing.Offices.Conversions.ICellValueConverter.GetStringValue|instance|System.String|System.Object|generic=0
- type|Bing.Offices.Conversions.ICellValueConverter|generic=0
```

`Bing.Offices.Core`：旧成员 `334`，新成员 `306`；新 hash：
`410F6A0F6CF64B41C3AB141AECFB2E1606C9B116EF2BD56C9A868FE02DC8FB68`。

```text
- constructor|Bing.Offices.Attributes.DateTimeAttribute..ctor||generic=0
- constructor|Bing.Offices.Attributes.DuplicationAttribute..ctor||generic=0
- constructor|Bing.Offices.Attributes.MaxLengthAttribute..ctor|System.Int32|generic=0
- constructor|Bing.Offices.Attributes.RangeAttribute..ctor|System.Int32,System.Int32|generic=0
- constructor|Bing.Offices.Attributes.RegexAttribute..ctor|System.String|generic=0
- constructor|Bing.Offices.Attributes.RequiredAttribute..ctor||generic=0
- field|Bing.Offices.CsvHelper.CsvQuoteCharacter|System.Char
- method|Bing.Offices.CsvHelper.GetCsvText|static|System.String|System.Data.DataTable,System.Boolean|generic=0
- method|Bing.Offices.CsvHelper.ToCsvBytes|static|System.Byte[]|System.Data.DataTable,System.Boolean|generic=0
- method|Bing.Offices.CsvHelper.ToCsvBytes|static|System.Byte[]|System.Data.DataTable|generic=0
- method|Bing.Offices.CsvHelper.ToCsvFile|static|System.Boolean|System.Data.DataTable,System.String,System.Boolean|generic=0
- method|Bing.Offices.CsvHelper.ToCsvFile|static|System.Boolean|System.Data.DataTable,System.String|generic=0
- property|Bing.Offices.Attributes.DateTimeAttribute.ErrorMsg|System.String
- property|Bing.Offices.Attributes.DuplicationAttribute.ErrorMsg|System.String
- property|Bing.Offices.Attributes.MaxLengthAttribute.ErrorMsg|System.String
- property|Bing.Offices.Attributes.MaxLengthAttribute.MaxLength|System.Int32
- property|Bing.Offices.Attributes.RangeAttribute.ErrorMsg|System.String
- property|Bing.Offices.Attributes.RangeAttribute.Max|System.Int32
- property|Bing.Offices.Attributes.RangeAttribute.Min|System.Int32
- property|Bing.Offices.Attributes.RegexAttribute.RegexString|System.String
- property|Bing.Offices.Attributes.RequiredAttribute.ErrorMsg|System.String
- property|Bing.Offices.CsvHelper.CsvSeparatorCharacter|System.Char
- type|Bing.Offices.Attributes.DateTimeAttribute|generic=0
- type|Bing.Offices.Attributes.DuplicationAttribute|generic=0
- type|Bing.Offices.Attributes.MaxLengthAttribute|generic=0
- type|Bing.Offices.Attributes.RangeAttribute|generic=0
- type|Bing.Offices.Attributes.RegexAttribute|generic=0
- type|Bing.Offices.Attributes.RequiredAttribute|generic=0
```

`Bing.Offices.Npoi`：旧成员 `2`，新成员 `2`；本轮无新增成员差异。旧 converter 构造参数属于非导出/internal 影响，不计入当前公开成员快照。

### 正式基线裁决

- `ICellValueConverter`、六个旧 validation attributes、`CsvSeparatorCharacter`、`CsvQuoteCharacter` 和五个旧 DataTable 重载均已从当前源码/API 中删除；显式 delimiter/quote 的 DataTable API 保留。
- Round 5 net6/net8 Unit 各为 `382 total / 381 passed / 1 failed`，唯一失败为正式 API hash mismatch；这是正式 baseline 获批前的历史结果。
- 正式 baseline 更新条件已在本执行回合满足：用户批准本轮 breaking change，已保存当前 Release 快照和完整 member diff，并已重跑 API compare 与所有可用 TFM Unit。

## Round 5 包产物身份

本轮 `dotnet pack` 输出目录为 `artifacts/packages-round5`。业务包 SHA-256：

| 包 | 版本 | SHA-256 |
| --- | --- | --- |
| `Bing.Offices.Abstractions` | `2.0.0` | `9911681786D724482F6E19DD08F9526463CC25759DD49E399247D99AF7FB03D0` |
| `Bing.Offices.Core` | `2.0.0` | `541D39F08C8F18A3CFFB69B0F6EF64CA375CA891702AE0F927A53D6272B2A2A2` |
| `Bing.Offices.Npoi` | `2.0.0` | `67932EEB37001A3B5099D2A52FF481FC2C78B8D28A6B8AD6FFD1E2C646C42483` |
