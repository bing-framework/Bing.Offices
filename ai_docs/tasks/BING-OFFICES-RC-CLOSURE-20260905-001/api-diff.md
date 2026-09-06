# Public API Candidate Diff

## 状态

`BLOCKED / formal approval pending`。候选成员快照已生成，但仓库中不存在历史 formal baseline JSON，只有 `PublicApiContractTest` 内的旧 hash，因此不能伪造逐行 formal diff，也未更新正式 hash。

## 候选快照

三 TFM（netcoreapp3.1/net6/net8）成员形态和 hash 一致：

| Assembly | Member count | Candidate hash | Formal hash |
| --- | ---: | --- | --- |
| Bing.Offices.Abstractions | 775 | `5458340E889F41ADBA12D1B76A21A0ADA5EC4F4025E804D795CAA6BFD3CCDEEF` | `7F9A2AA819E94B3838097DF2FF374A934CF7F35F3D2E91F3D1DB790F22972943` |
| Bing.Offices.Core | 152 | `F2C082AFD032D59FE8AA0F0DD54A585A45611DF72F4157E898D33AC1B8D5BE98` | `B3661970BBE5AECC06DAD57B1E3F960FA77E70C4D2E66B2DA4910F7823AA2BB6` |
| Bing.Offices.Npoi | 66 | `1C5E5DCEA0BDDA945ABCEED01B4DFFE24EE98593032D306C86FF4E61FECC7427` | `DA163263804A964D8AC2A13D78D6B3858256171CE7729841690FDB56F602CEEE` |

原始成员清单：`artifacts/api-snapshot/candidate/api-snapshot-*.json`。

## 本任务明确变化

- 删除 `ExcelMapping`、`ExcelMappingBuilder<T>`、`ExcelColumnMappingBuilder<T,TProperty>`。
- 删除 DataTable `Bing.Offices.CsvHelper`。
- 删除 runtime `MigrateV1Json/Xml` 全部重载。
- 删除 `ExpressionExtension`、`RegexConst`、`ExcelValueMap<T>`；`PropertyInfoExtensions`、`TypeExtensions` 和六个默认 validation 实现不再导出。
- `ExportColumnMappingBuilder<T,TProperty>` 新增 `HasConverter` 与 `Map`，承接旧 Builder 的真实导出能力。
- `ExcelResourceLimits.MaxInputBytes` 默认值改为 128 MiB，不改变成员签名。
- `BingOfficesException` 改为 abstract，保留统一 catch 与派生异常合同；当前快照 canonicalizer 不编码类型 abstract 修饰符，因此 hash 不变，直接契约测试补充锁定该变化。
- 七个 NPOI 扩展容器全部保留 public。

最终候选重新生成到 `artifacts/api-snapshot/candidate-final`；三 TFM 的成员数与 hash 仍一致。正式 baseline 仍因缺少历史成员 JSON 而 `BLOCKED`。

## 工具与复现

ApiSnapshot CLI 新增 `--dependencies`，用于沙箱/CI 显式提供元数据依赖目录。候选采集使用空 seed 仅驱动输出，不代表正式批准 baseline。

```powershell
dotnet build build/ApiSnapshot/ApiSnapshot.csproj -c Release --no-restore
dotnet build/ApiSnapshot/bin/Release/net8.0/ApiSnapshot.dll --root output/release --baseline ai_docs/tasks/BING-OFFICES-RC-CLOSURE-20260905-001/artifacts/api-snapshot/capture-seed.json --dependencies tests/Bing.Offices.Tests/bin/Release --output ai_docs/tasks/BING-OFFICES-RC-CLOSURE-20260905-001/artifacts/api-snapshot/candidate
```
