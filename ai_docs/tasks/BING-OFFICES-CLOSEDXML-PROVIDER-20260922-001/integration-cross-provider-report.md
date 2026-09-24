# Cross-provider Integration Report

## Status

`PASS_WITH_BLOCKED_APPROVAL`：`tests/Bing.Offices.Tests.Integration` 在 net6.0 和 net8.0 各有 38 个测试通过；另有 1 个 Public API Snapshot 测试因批准 baseline 缺少 ClosedXML assembly 而按预期 `BLOCKED_APPROVAL`。新增合同使用同一公共 Request 对 NPOI、MiniExcel、ClosedXML 执行基础标量、Mapping、动态列、Relations、Validation、RowHeight 和资源限制往返，并对 NPOI/ClosedXML 验证公式数值/布尔缓存值及字符串公式策略；同时验证 ClosedXML 公共类型不暴露 `ClosedXML.Excel` DOM 类型。

执行命令：

```text
dotnet restore tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj --disable-parallel --ignore-failed-sources -p:NuGetAudit=false --packages C:\Users\jianx\.nuget\packages
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj --no-restore -c Release -v:minimal
```

本结果仍建立在当前 dirty worktree 上；它是结构化基础合同证据，不等同于所有 Rich XLSX 能力已跨 Provider 统一。

## Required follow-up

- Cross-provider 共同请求已扩展到 Mapping、Formula、RowHeight、Relations、Validation、Dynamic Column 和资源限制；剩余 Provider-specific Unsupported 行为和 first-registration-wins 证据已由各 Provider 直接测试覆盖。
- 在批准 API snapshot 流程中加入 ClosedXML 程序集的成员级快照；本地 capture 和 `api-diff.md` 已完成，baseline 更新仍需维护者批准。
