# Unit Test 报告

## 结论

状态：`FAILED`。三个 TFM 的行为测试均通过，唯一失败是执行前已存在且本轮必须保留的正式 Public API hash 门禁。没有跳过测试，没有删除或放宽失败断言。

## 环境与身份

- Git：`master` / `9d78ab76e28891ac9b7e0e558f4df669721a4ea6` / dirty（执行前 clean）。
- Windows 10.0.19045 win-x64；SDK 10.0.400。
- Runtime：.NET Core 3.1.32、.NET 6.0.36、.NET 8.0.30。
- 配置：Release；项目 `tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj`。

## 最终结果

| TFM | Discovered | Passed | Failed | Skipped | 耗时 | Artifact |
| --- | ---: | ---: | ---: | ---: | ---: | --- |
| netcoreapp3.1 | 471 | 470 | 1 | 0 | 约 2 s | `artifacts/test-results/unit/rc-final-unit-netcoreapp3.1.trx` |
| net6.0 | 471 | 470 | 1 | 0 | 约 2 s | `artifacts/test-results/unit/rc-final-unit-net6.0.trx` |
| net8.0 | 471 | 470 | 1 | 0 | 约 3 s | `artifacts/test-results/unit/rc-final-unit-net8.0.trx` |

三次唯一失败均为 `PublicApiContractTest.PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`：Abstractions expected `7F9A2AA819E94B3838097DF2FF374A934CF7F35F3D2E91F3D1DB790F22972943`，actual `5458340E889F41ADBA12D1B76A21A0ADA5EC4F4025E804D795CAA6BFD3CCDEEF`。正式成员 baseline JSON 不存在，无法执行成员级审批，因此没有更新内嵌 hash。

## 命令

```powershell
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f <TFM> --no-build --no-restore --logger "trx;LogFileName=final-unit-<TFM>.trx" --results-directory <task>/artifacts/test-results/unit
```

专项结果包括：关系缓存/异常 18/18；输入与 ZIP 资源矩阵 31/31；Atomic/File 9/9；图片 HSSF/XSSF 与未知 Sheet；DateTimeOffset 在 `TZ=UTC` 与 `TZ=Pacific Standard Time` 下分别 6/6；配置 Loader 4/4；Failure Workbook 无 sink/sink 失败 Trace 2/2。

## 覆盖率与风险

项目未引用 coverage collector，本轮未伪造覆盖率百分比。原始证据为 TRX。剩余风险是正式 API 成员基线缺失；netcoreapp3.1 已 EOL，且若干 8.0 依赖不声明支持该 TFM。
