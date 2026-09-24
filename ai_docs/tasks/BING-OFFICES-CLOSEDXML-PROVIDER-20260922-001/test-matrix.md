# Test Matrix

| Level | Command | Result |
| --- | --- | --- |
| L0 | `git diff --check` | PASS（现有 dirty worktree 警告保留） |
| L0 | ClosedXML project build `--no-restore` | PASS, net6/net8 |
| L1/L2 | `dotnet test tests/Bing.Offices.ClosedXml.Tests/...csproj --no-restore` | PASS, 74 tests × 2 TFM；覆盖固定列混合显式索引、动态列稀疏物理布局/冲突/列宽、列级样式/行高、导入边界/Error Cell、Mapping Plan 类型/配置/动态隔离与失败恢复、重复表头、MaxSheets/MaxColumns/MaxCells、准入队列生命周期、Entity Relations/负向和原子提交清理 |
| L3 | `dotnet test tests/Bing.Offices.ClosedXml.Tests.Integration/...csproj --no-restore` | PASS, 4 tests × 2 TFM；含同步文件替换/失败保留和真实读回 |
| L4 | `dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj --no-restore` | ClosedXML Mapping/Formula cached value/动态列/Relations/RowHeight/Validation/资源限制合同和公共 API 分类共 38/39 × 2 TFM 通过；每个 TFM 唯一失败为批准 baseline 尚未纳入 ClosedXML/新增成员，保留 `BLOCKED_APPROVAL` |
| L0 | Benchmark project Release build `--no-restore` | PASS, net8；显式 ClosedXML 运行时依赖已验证 |
| L4 | `dotnet test Bing.Offices.sln --no-restore -c Release -v:minimal -m:1` | 共 1944 tests，成功 1942、失败 2；Core 201×2、MiniExcel 43×2 + Integration 9×2、NPOI 567×2 + Integration 30×2、ClosedXML 74×2 + Integration 4×2、Docs 10 通过；公共 Integration 38/39×2，两个失败均为未批准 API snapshot，`BLOCKED_APPROVAL` |
| L3 | Package Consumer net6/net8 + ThirdParty Provider Consumer | PASS；任务专属本地 nupkg、无 ProjectReference，net6/net8 实际 ClosedXML export/import 读回通过 |
| L5 | 1K/10K/100K ClosedXML Probe | PASS；普通 round-trip 与独立 export/import/file/multisheet/style/template/formula 场景均以 sync/外围 async 各 3 次运行，原始 JSONL 和中位数见 `benchmark-report.md` |
| L5 | 500K/1M、跨 OS/字体、完整大容量 Resource Matrix | `BLOCKED_APPROVAL` / `BLOCKED_EXTERNAL`；未伪造性能或容量数字 |

NuGet restore initially遇到 `NU1301`，经一次受控 restore 后成功完成；依赖文件没有手工修改。
