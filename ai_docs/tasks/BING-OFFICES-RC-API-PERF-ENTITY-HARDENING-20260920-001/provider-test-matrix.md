# Provider 测试矩阵

| Provider | Unit net6 | Unit net8 | Integration net6 | Integration net8 | 备注 |
|---|---:|---:|---:|---:|---|
| NPOI | 566/566 | 566/566 | 30/30 | 30/30 | Entity 专项双 TFM 已通过；固定 Cell、Merge、List、Template、HSSF、Async、LeaveOpen、样式/图片/复杂 Merge、边界、命名转换器、文件取消保护、非 `IList` 关系集合、100K 流式列表、属性级 Merge 回归和大型列表取消保护通过 |
| MiniExcel | 42/42 | 42/42 | 9/9 | 9/9 | 日期、RawDate、dynamic、ICollection、非 `IList` relation、Entity fail-fast 通过 |

公共集成项目 `Bing.Offices.Tests.Integration`：当前两 TFM 均 `39/39 PASS`；早期 Round 3 的 `27/29` 仅保留为历史 API baseline/snapshot 失败记录，不是当前 Provider 运行时失败。公开包第三方 fixture net6/net8 build/run 均通过。
