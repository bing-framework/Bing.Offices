# 单元测试报告

## 执行环境

- OS：Windows 10.0.19045 x64
- SDK：.NET SDK 10.0.400
- 可用运行时：net6.0、net8.0
- 缺失运行时：netcoreapp3.1、net5.0、net7.0
- 配置：Release，`--no-restore`；结果保存为 TRX

## 正式 API baseline 批次

用户在本执行回合明确批准将当前 Release API 纳入正式 baseline。正式 baseline：`artifacts/api-snapshot-formal-baseline-20260903.json`；对应当前 Release 快照：`artifacts/api-snapshot-formal-20260903/api-snapshot-*.json`。API compare 退出码为 `0`，三个逻辑程序集在 `netcoreapp3.1`、`net6.0`、`net8.0` 快照中的成员数/hash 一致。

| TFM | 命令目标 | 总数 | 通过 | 失败 | 跳过 | 退出码 | TRX |
| --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| net8.0 | `tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj` | 384 | 384 | 0 | 0 | 0 | `tests/Bing.Offices.Tests/TestResults/api-baseline-net8-final-rerun.trx` |
| net6.0 | `tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj` | 384 | 384 | 0 | 0 | 0 | `tests/Bing.Offices.Tests/TestResults/api-baseline-net6-final-rerun.trx` |

正式 API hash：Abstractions `7F9A2AA819E94B3838097DF2FF374A934CF7F35F3D2E91F3D1DB790F22972943`；Core `B3661970BBE5AECC06DAD57B1E3F960FA77E70C4D2E66B2DA4910F7823AA2BB6`；NPOI `DA163263804A964D8AC2A13D78D6B3858256171CE7729841690FDB56F602CEEE`。

## 历史基线未批准批次

以下“唯一失败”记录属于正式 baseline 获批前的历史执行，不代表当前结果；本节保留以证明当时未通过修改断言掩盖差异。

## 已执行结果（历史）

| TFM | 命令目标 | 总数 | 通过 | 失败 | 跳过/未执行 | 退出码 | TRX |
| --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| net8.0 | `tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj` | 382 | 381 | 1 | 0 | 1 | `tests/Bing.Offices.Tests/TestResults/unit-net8-review-fix-round1.trx` |
| net6.0 | `tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj` | 383 | 382 | 1 | 0 | 1 | `tests/Bing.Offices.Tests/TestResults/unit-net6-review-fix-round1.trx` |

命令模板：

```powershell
 dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -c Release --no-restore -f net8.0 --logger "trx;LogFileName=unit-net8-review-fix-round1.trx"
 dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -c Release --no-restore -f net6.0 --logger "trx;LogFileName=unit-net6-review-fix-round1.trx"
```

## 失败详情（历史）

两个可运行 TFM 的唯一失败均为：

`PublicApiContractTest.PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`

失败内容是测试内固定的旧 hash 与当前实际 API 不一致：

- 期望 Abstractions：`7B0BA2792AE1DB91BB281C1719B0B35671091CA981C659FDC89B3771B7F5F577`
- 实际 Abstractions：`407B4F3C2605333A082766B13E1F1DEB704880DDE6D3E0CEDB72FF3F6281ADF0`
- 期望 Core：`41B6D12CD58A988E84701902E0F58476B33903583A51F39DF7544B436504DF54`
- 实际 Core：`5F68499B76921FA52D293BC851B94659FBB2AB466E6D91C8878A97F8728B4BB7`
- 期望 NPOI：`A0DBE9808D82547601429D8958C7ED283467031A3763EB9037B19D03F19D80BD`
- 实际 NPOI：`DA163263804A964D8AC2A13D78D6B3858256171CE7729841690FDB56F602CEEE`

逐成员比较表明：Core 与既有 `api-snapshot-review10` 完全一致；Abstractions 只有 `MaxBytes -> MaxSerializedBytes` 一项成员替换；NPOI 只有 `AddNpoi -> AddBingOfficesNpoi` 一项成员替换。该失败是未裁决的 API 基线更新问题，不是通过降低断言解决的问题。`5A1B...` 是旧 snapshot 的值，不是本轮 Release 的实际 Abstractions hash。

## 定向回归

新 DI 入口相关定向测试：`6/6` 通过，覆盖 null 校验、链式返回、注册、替换和服务解析。NPOI exact member baseline 测试通过。

## 覆盖分类

- Mapping request/profile/document 合并、方向迁移和缓存隔离：有回归用例并在上述运行中执行。
- 关系委托异常类型保持：五类委托场景有回归用例并通过。
- CSV 公式前缀、`None` 策略、非法 options、RFC4180 和 culture：有回归用例并通过。
- Failure Workbook、原子提交、取消、流所有权和清理诊断：有回归用例并通过。
- Excel ResourceProbe：独立结果见 `resource-report.md` 的七模式 child-process artifact；mapping/unique ResourceProbe 为另一套独立 JSONL，不能混作 Excel 资源证据。

## Package consumer 交叉证据

`artifacts/package-consumer-rerun2` 是无 `ProjectReference` 的独立 `PackageReference` consumer。使用最新本地 2.0.0 nupkg 和短路径 `C:\nupkg-cache` 时 restore/build/run 均退出码 `0`，输出 `package-consumer-ok`；任务深路径缓存下 build 因 SDK/MSBuild `MSB3106` 失败。该证据仅覆盖 net8.0 consumer，不改变本报告的 TFM 覆盖范围和其它 package/交付限制。

## 环境限制

netcoreapp3.1、net5.0、net7.0 testhost 需要的运行时未安装。本报告不把这些 TFM 写成通过；解除条件是安装对应 runtime 后按相同命令重跑。

## 状态

`PARTIAL`：正式 API baseline 批次的可运行 TFM Unit 已全绿；netcoreapp3.1、net5.0、net7.0 runtime 仍未安装，相关 TFM 未运行。

## Round 5 最新全量回归

本节以 Round 5 删除已迁移 `[Obsolete]` API 后重新执行的 TRX 为当前结果；上文 Round 1/历史结果保留为历史证据，不覆盖。

| TFM | 总数 | 通过 | 失败 | 跳过 | 退出码 | TRX |
| --- | ---: | ---: | ---: | ---: | ---: | --- |
| net8.0 | 382 | 381 | 1 | 0 | 1 | `tests/Bing.Offices.Tests/TestResults/review-fix-round5-unit-net8.trx` |
| net6.0 | 382 | 381 | 1 | 0 | 1 | `tests/Bing.Offices.Tests/TestResults/review-fix-round5-unit-net6.trx` |

两个 TFM 的唯一失败仍为 `PublicApiContractTest.PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`，均为正式 baseline 获批前的历史 hash mismatch；其余行为测试通过。Round 5 定向 CSV/校验回归为 `6/6`，Integration net6/net8 为 `30/30`，Docs net8 为 `11/11`，详见对应报告。
