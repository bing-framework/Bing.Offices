# Package Consumer Report

- Task ID：`BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001`
- 包目录：`artifacts/packages/`
- Consumer 目录：`artifacts/consumers/`

## 包检查

最终 pack 产物为四个 `2.0.0` 版本的 `.nupkg` 和四个 `.snupkg`，均位于 `artifacts/packages/`。本轮最终 `.nupkg` SHA-256 为：Abstractions `A177B00739E2A2986DC07A1A1A1B90EE36C57457300CEEA4F55CF2B1B191BCB2`、Core `F4A35810496780F784808F523D61DF722B38910B3AEBF3AF3BF6318EC5BB4ADF`、Npoi `CCA607AA9CBDBCDF0CCBAACF975274F082D15B54A4E46CFB1B1391A195797BD1`、MiniExcel `36F1DFDBC029925E50DAF36249E79554F50A974F0F6858FE1EC063AABC30C32F`。每个包均包含对应 `lib` DLL/XML；MiniExcel nuspec 仅声明 Core、DI Abstractions 和 MiniExcel 依赖，没有 NPOI 依赖。完整检查结果见 `artifacts/packages/package-check.txt`。

最终包内 DLL 与 Consumer 输出的 SHA-256 逐项相符：net6 的 Core `41F03B6A3DBF7655927780ECA94485CB0605261FE36B670CA1A9E66FF1CDE7C0`、Npoi `C87FCE314EA7B9A828264EBD0E6DBB383D4CE815B99D5380362735DDE30F3BD8`、MiniExcel `D9D12BF67D5E761DE88D95BF58D574B06F86B57286FF41D00054EF8C4948BC24`；net8 的 Core `41F03B6A3DBF7655927780ECA94485CB0605261FE36B670CA1A9E66FF1CDE7C0`、Npoi `94E75DC32C73F923ED9B0985EA8F61F75C7023898043711BA25F2C479AF84BF1`、MiniExcel `E523622EF2F9A9C0C4135341FF64DF838FA87E9351AA3B12C8C110E59E1938B4`。

## PackageReference 消费

Net6 与 net8 Consumer 均保留 `PackageReference`，由 `BingOfficesPackageVersion=2.0.0` 注入本次 `artifacts/packages` 中唯一发现的实际 nuspec 版本，并在独立 cache 中 restore 后运行真实 NPOI/MiniExcel API。最终包写入时间早于本轮 Consumer 日志（包约 `08:39:50Z`-`08:39:56Z`，日志约 `08:45:59Z`-`08:46:01Z`）：

- net6 restore/build/run：通过；独立 cache `artifacts/consumers/review-round2-final/cache/net6/`，日志 `artifacts/consumers/review-round2-final/net6-build.log`、`net6-run.log`
- net8 restore/build/run：通过；独立 cache `artifacts/consumers/review-round2-final/cache/net8/`，日志 `artifacts/consumers/review-round2-final/net8-build.log`、`net8-run.log`
- 输出均包含 `package-consumer-ok` 和 `Bing.Offices.Npoi+Bing.Offices.MiniExcel/2.0.0`

## 限制

本轮新的独立 cache restore 已完成；此前 `artifacts/consumers/review-round2/` 的空 cache TLS 失败日志保留为历史环境限制，不再作为本轮最终包证据。外部 CI、生产环境和 500K/1M 容量验证仍按计划为 `NOT_VERIFIED`。
