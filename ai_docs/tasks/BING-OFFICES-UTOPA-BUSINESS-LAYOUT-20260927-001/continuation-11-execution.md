# Continuation 11：最终 Footer 命名锚点执行记录

## 结果

本阶段 COMPLETED。`FooterNamed` 把工作簿名称定位到本次最终 Footer marker；空明细、变长明细和带分页小计的布局都随实际行数移动。模板中同名、同 Sheet 的单格名称会更新，缺失时创建；跨 Sheet、范围或歧义名称返回 Plan 配置错误。导入仍按 marker 文本判断明细边界，旧 `Footer` 行为保持不变。

NPOI XLS/HSSF、XLSX/XSSF 与 ClosedXML XLSX 均实现该行为。PageSubtotal 和 GroupSubtotal 可重复出现，仍不绑定单格名称。

## 验证证据

| 层级 | 验证 | 结果 |
| --- | --- | --- |
| L1 | `FooterNamed`、模板名称、冲突和取消，NPOI | net6.0、net8.0 各 10/10 |
| L1 | 同上，ClosedXML | net6.0、net8.0 各 6/6 |
| L2 | NPOI 常规测试，`Category!=Large` | net6.0、net8.0 各 685/685 |
| L2 | ClosedXML 常规测试，`Category!=Large` | net6.0、net8.0 各 202/202 |
| L3 | Provider Contract，`Category!=Large` | net6.0、net8.0 各 282/282 |
| L3 | Public API Integration | net6.0、net8.0 各 11/11 |
| L3 | Docs 可执行示例 | net8.0 10/10 |
| L0 | Solution Release build | 0 错误；4 个既有 Microsoft.Bcl.Memory net6.0 警告 |
| L0 | Solution 常规测试，`Category!=Large`，单项目顺序运行 | 全部通过，共 3168 个测试实例 |

`dotnet pack` 后，第三方消费者使用本轮 `output/release` NuGet 包；net6.0、net8.0 均构建并输出 `third-party-public-only-provider-ok`。首次命令行包源参数使 NuGet 将 HTTPS 地址解析为本地路径；改用阶段专属 `footer-anchor-nuget.config` 后还原通过。`UseAppHost=false` 构建后直接运行 DLL，两个 TFM 均通过。最终配置改为相对包源加 nuget.org，再次还原通过；离线环境只出现 NuGet 漏洞数据读取警告。

API 更新前的 `api-footer-anchor/capture/api-candidate.json` 与原基线逐成员比较：net6.0、net8.0 均只新增 `FooterNamed` 方法和隐藏的 `FooterAnchorName` 属性，无成员删除或签名变化。成员级批准记于 `api-approval.md`；批准后重新捕获身份并更新双 TFM 基线，`api-footer-anchor/compare-final/api-diff.json` 两框架均为 `{}`。未更改版本号。

`git diff --check` 无空白错误。严格 UTF-8 字节检查确认本阶段 C# 为 BOM+LF，Markdown、验证配置为无 BOM+LF，均有末尾换行。`build/api-snapshot-baseline.json` 延续前阶段已有 CRLF、无末尾换行格式，`git` 因其 `eol=lf` 提示规范化；本阶段未混入整仓编码转换。`ProfileFixtures.xml` 也保留原有 CRLF 提示。外部 Utopa 仓库未改动，未提交、推送、创建 PR 或发布。

## TODO 分类

- Completed：最终 Footer 命名锚点、两 Provider 真工作簿验证、双 TFM API/包消费者、文档与追溯。
- Open Actionable：无。
- Blocked Approval / Blocked External：无。
- Accepted Limitations：一个名称只指向最终 Footer 单格；重复的分页/分组小计不绑定名称，导入不依赖名称元数据。
- Deferred：可复用 Entity Layout Provider 合同模板、Roslyn 地址/动态 Key 分析器；只有出现启动性能证据后再评估 Source Generator。
- Next Action：本阶段停止，后续优先设计可复用 Provider 合同模板，再评估分析器。

IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
REVIEW_STATUS: NARROW_SCOPE_CHECKED
RELEASE_STATUS: NOT_EVALUATED
OPEN_ACTIONABLE: 0
