# 单元测试报告

| 项目 | TFM | 结果 |
|---|---|---|
| `Bing.Offices.Tests` | net6.0 | `197/197 PASS` |
| `Bing.Offices.Tests` | net8.0 | `197/197 PASS` |
| `Bing.Offices.Npoi.Tests` | net6.0 | `566/566 PASS`；Entity 专项 `22/22 PASS` |
| `Bing.Offices.Npoi.Tests` | net8.0 | `566/566 PASS`；Entity 专项 `22/22 PASS` |
| `Bing.Offices.MiniExcel.Tests` | net6.0 | `42/42 PASS`；Entity fail-fast `2/2 PASS` |
| `Bing.Offices.MiniExcel.Tests` | net8.0 | `42/42 PASS`；Entity fail-fast `2/2 PASS` |
| `Bing.Offices.Docs.Tests` | net8.0 | `10/10 PASS` |

专项证据：MiniExcel 关系过滤 `6/6 PASS`（含非 `IList` `ICollection<T>`）；NPOI 新增关系/只读/图片筛选 `4/4 PASS`，Entity/Template 样式、图片锚点、非精确 Merge、10,000 行边界和文件取消保护职责级测试双 TFM `22/22 PASS`；新增 100,000 行 SXSSF 流式列表 round-trip/样式测试双 TFM 各 `1/1 PASS`；第三方 public-only fixture net6/net8 均输出 `third-party-public-only-provider-ok`；TRX 位于 `artifacts/tests/`。测试使用隔离 NuGet 缓存编译并引用同一次 Release build 的输出。

注：早期快照中的 NPOI `559/559` 与 Entity `15/15` 已由当前候选回归扩展为上表的 `566/566` 与 `22/22`，不作为当前结果使用。
