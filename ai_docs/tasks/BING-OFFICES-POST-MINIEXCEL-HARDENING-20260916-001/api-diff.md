# API Diff

- Task ID：`BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001`
- 分支：`feat/miniexcel-provider`
- 比较范围：Release 双 TFM（`net6.0`、`net8.0`）以及四个 NuGet 包。

## 结论

本任务没有新增、删除或修改公共成员。最终 `api-diff.json` 为空，双 TFM snapshot compare 通过。MiniExcel 的公开 Provider 类型已经包含在快照治理范围内，但其公开签名来自前一任务；本任务没有新增 Provider SPI，也没有删除既有公共成员。

生产程序集之间的 `InternalsVisibleTo` 已清理为 0；仅保留测试程序集友元。Core 的预检实现通过共享源链接编译到 NPOI 和 MiniExcel Provider，避免扩大公共 API。

## 证据

- API capture：`artifacts/api-round2-final3/capture/`
- API compare：`artifacts/api-round2-final3/compare/`
- Compare 日志：`artifacts/api-round2-final3/compare.log`
- Identity self-test：`artifacts/api-round2-final3/identity-self-test.log`
- 包内 DLL 篡改验证：`artifacts/api-round2-final3/compare-tampered.log`（预期退出码 `1`）
- 最终结果：`API snapshot comparison passed for net6.0 and net8.0.`
- Identity 结果：LF/CRLF clean checkout、source/assembly/nupkg/approval tamper 检查均按预期通过或失败；实际包内 MiniExcel DLL 替换为文本后 compare 以退出码 `1` 拒绝。

## 元数据

候选源码清单哈希已按最终源码重新生成：`34B821D2678CB33FCE7DC26A4D26BDCA099636E699B635860BC94A16BF1DD29C`。程序集 identity 与公共成员哈希未发生 API 变化；baseline 中只同步了可审计的候选源码和包 identity 元数据，没有自动批准新的公共 API。
