# 当前候选基线

## Candidate

- Task ID: `BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001`
- Base commit: `94bb52e84ffc70634067b04541857433ef7af9df`
- Branch: `feat/miniexcel-provider`
- Initial candidate source manifest SHA-256: `A81EE028043855611280D3D6AE02A03A642C778A35104E5EBD1EA722BD71C007`
- Worktree state: `dirty`（当前任务变更及任务文档存在）
- Runtime: Windows `10.0.22631`, `win-x64`
- SDK: `.NET SDK 10.0.401`
- Tested runtimes: `.NET 6.0.36`, `.NET 8.0.30`
- Configuration: `Release`

## Immutable Hashes

| 文件 | SHA-256 |
|---|---|
| `version.props` | `EC5201C9A9EC7F4981879FFA604C5BA2B35F89F064AB88C790ECFB6644DC5306` |
| `version.dev.props` | `0A2F75F575DCC863D51F256A0D823F400FC33369FB1C7EB823CE6EBA9321872B` |
| `common.props` | `1C830F653129C28F59094F3F5886EB1BA1640ACBFE8984FF34F24FEDAB10ECA4` |
| `framework.props` | `5B64D45A8FC516449251449C619E8F0434481A295D18A8B727AC4978FF5FDE35` |
| `build/api-snapshot-baseline.json` | `951F9944C39B843CFDFE8DF92BA7C4592F4A9CAB2F459540E4FA6B7EF1F4AAF0` |

## Baseline Limitations

初始 restore 在受限沙箱中因 NuGet SSL/凭据返回 `NU1301`，获得外部 restore 权限后才完成依赖恢复。因此本任务没有可信的 clean-before 六项目双 TFM 测试或 before benchmark；所有 before 对比均标记 `NOT_AVAILABLE_BEFORE`，不能从历史产物推导提升比例。

当前候选的 Release solution build、双 TFM职责测试、包、Consumer、API capture 和 100K after-only 探针均使用本报告的 base commit 与 dirty source manifest。历史无限制 500K/1M 不作为容量证据；本轮 Job Object 受限 500K/1M 结果见 `resource-report.md`，生产机器和外部 CI 仍未执行。

本报告记录的是执行早期的 baseline 采样；早期候选身份保留在 `manifest.md` 和 `artifacts/api/current/api-candidate.json`，其 `candidateSourceManifestSha256` 为 `22F07B691006F23F9EDA061E1FE709751199023DCCD7BA5A4C3C0259EABD488A`。随后基于当前工作树形成的身份刷新记录为 `manifest-current-20260921.md` 和 `artifacts/api/current-20260921-candidate/api-candidate.json`，其 source hash 为 `E4447DEB168499218EE856357197D80A96208F625F79F4DCABDCB1E4AB101346`。两者的 base commit、版本文件和 API baseline 不变；早期 hash 不用于当前刷新候选身份，旧 benchmark/resource 证据也不自动迁移。
