# PackageReference Consumer 报告

生产包从当前 Release 输出重新生成至 `artifacts/packages`，版本读取为既有 `2.0.0`；版本文件哈希未变化。NPOI 流式路径修复后，最终 Consumer 使用 `artifacts/consumer/NuGet.Config` 的本地包源和官方依赖源，并显式指定全新 `--packages artifacts/consumer/cache-rc-final7-*`，未修改项目版本或混用旧同号缓存。

| Consumer | Restore | Build | Run |
|---|---|---|---|
| net6.0 | PASS（全新隔离缓存；外部 NuGet restore） | PASS（1 个既有 `NETSDK1138` EOL warning） | PASS，`package-consumer-ok` |
| net8.0 | PASS（全新隔离缓存；外部 NuGet restore） | PASS | PASS，`package-consumer-ok` |

运行输出包含：`package-consumer-ok`、NPOI/MiniExcel 包版本 `2.0.0`、CSV/Excel sync/async 往返和 Provider 扩展验证；当前重打包后 net6 输出 `csvBytes=31`、`excelBytes=4272`、`miniExcelBytes=4210`，net8 输出 `csvBytes=31`、`excelBytes=4272`、`miniExcelBytes=4209`。首次受限沙箱 restore 的 `NU1301` 已通过批准的外部 restore 重试解决；本轮使用全新隔离缓存 `artifacts/consumer/cache-rc-final7-net6` 与 `artifacts/consumer/cache-rc-final7-net8` 复核。

独立 public-only Provider fixture `tests/Bing.Offices.ThirdPartyProvider.Consumer` 只引用四个已打包程序集，同样使用全新隔离缓存 `artifacts/consumer/cache-rc-final7-public-net6` 与 `artifacts/consumer/cache-rc-final7-public-net8`，net6/net8 build/run 均 PASS；输出 `third-party-public-only-provider-ok`，覆盖公开 Entity/Template SPI、能力 preflight、取消、流所有权、文件端点和结构化 unsupported 异常。此前旧缓存和用户全局同号包结果均已排除，不计入最终证据。当前 NPOI 包内 DLL 的物理 SHA-256 已同步到候选 manifest；其余三个包物理哈希保持不变。

## 当前候选身份刷新（2026-09-21）

基于当前工作树重新执行 Release build/pack，包输出到 `artifacts/packages/candidate-20260921`，API capture 输出到 `artifacts/api/current-20260921-candidate`。使用候选专用 NuGet 配置和全新缓存重新验证：

| Consumer | Restore | Build | Run |
|---|---|---|---|
| `Bing.Offices.Consumer.Net6` | PASS | PASS（1 个既有 `NETSDK1138` EOL warning） | PASS，`package-consumer-ok`，`csvBytes=31`、`excelBytes=4272`、`miniExcelBytes=4207` |
| `Bing.Offices.Consumer.Net8` | PASS | PASS | PASS，`package-consumer-ok`，`csvBytes=31`、`excelBytes=4272`、`miniExcelBytes=4209` |
| `Bing.Offices.ThirdPartyProvider.Consumer` net6 | PASS | PASS（1 个既有 `NETSDK1138` EOL warning） | PASS，`third-party-public-only-provider-ok` |
| `Bing.Offices.ThirdPartyProvider.Consumer` net8 | PASS | PASS | PASS，`third-party-public-only-provider-ok` |

本轮对应缓存为 `artifacts/consumer/cache-candidate-20260921-*`，当前 source-scope hash 为 `E4447DEB168499218EE856357197D80A96208F625F79F4DCABDCB1E4AB101346`。身份明细见 `artifacts/candidates/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/manifest-current-20260921.md`；该刷新解决当前 build/package/API capture 的内部一致性，但不把旧 benchmark/resource 产物静默归属于新候选。
