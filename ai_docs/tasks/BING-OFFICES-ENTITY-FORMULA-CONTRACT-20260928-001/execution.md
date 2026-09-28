<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-ENTITY-FORMULA-CONTRACT-20260928-001
AI_EXECUTION_FINISHED_AT: 2026-09-28T09:49:22Z

# 实施执行报告

## 执行结论

独立 Entity Provider 合同包新增第七项 `VerifyFooterFormulas`。合同通过 XLSX 工作簿关系定位工作表，以原生公式节点验证显式公式和连续明细求和；覆盖两条明细、空明细、`GapRows`、marker 与导入边界，并验证跨分页小计的最终连续求和在布局阶段拒绝。主任务当前阶段索引已更新至 Continuation 18。

## 任务信息

- 计划：`ai_docs/tasks/BING-OFFICES-ENTITY-FORMULA-CONTRACT-20260928-001/plan.md`。
- 范围：共享合同、官方合同入口、包消费者、接入文档、进度索引与追溯。

## 计划执行情况

三个实施项均已完成。合同包仍直接编译共享合同源码；第三方消费者从新打的包执行第七项合同。公式不在合同中求值，NPOI XLS/HSSF 专属行为仍由 Provider 自身测试负责。

## 已完成事项

- `VerifyFooterFormulas` 验证公式类型和完整公式文本、空/非空明细、原生 marker、导入往返与非法组合。
- 工作表解析使用 `workbook.xml` 与关系部件，允许第三方 Provider 使用不同的合法工作表文件名。
- 官方 NPOI/ClosedXML 合同测试和包消费者接入，文档与生产符号到测试方法映射同步。

## 部分/未完成事项

无本阶段遗留。跨中间小计的最终总计公式辅助、分页与分组小计组合等属于后续独立能力。

## 修改文件

合同源码 `EntityLayoutProviderContractSuite.cs`、合同测试 `EntityLayoutReusableContractTest.cs`、包消费者 `Program.cs`、`docs/excel/entity-provider-contracts.md`、Utopa 任务的 `execution.md` 和 `production-symbol-test-map.md`，以及本阶段计划与执行记录。

## API/数据/配置变化

合同测试包新增 `VerifyFooterFormulas(IExcelEntityExporter, IExcelEntityImporter)` 公开入口；八个生产程序集没有代码或公开 API 变化，未调整生产 API snapshot、包版本、数据或配置。

## 测试结果

- 新合同定向测试：NPOI、ClosedXML 在 net6.0、net8.0 各 2/2 通过。
- Provider Contract `Category!=Large`：net6.0 与 net8.0 各 297/297 通过。
- Docs：net8.0 10/10 通过。
- 新合同包 Release 打包成功；新隔离缓存的 `.nupkg.metadata` 指向本地 `artifacts/packages-contracts`。第三方包消费者 Release 构建通过，net6.0、net8.0 均输出 `third-party-public-only-provider-ok`。

## Build/Typecheck/Lint/Format

合同包与消费者构建 0 错误；修改文件严格 UTF-8、C# BOM+LF、Markdown 无 BOM+LF 和末尾换行检查通过；`git diff --check` 通过。未重复运行生产解决方案全量测试或 API 比较，因为本阶段仅改合同测试和文档，生产源码及快照范围未变。

## 计划偏差

初版合同断言了一个没有交给导出器的目标流，无法证明文件保护，因此删除该无效断言。合同现只检查布局阶段拒绝；实际目标流保护继续由 NPOI 和 ClosedXML 原有职责测试覆盖。合同按工作簿关系定位工作表，避免假定部件名为 `sheet1.xml`。

## 基线问题

构建输出仍有存量 net6.0 支持警告。`git diff --check` 输出的 API baseline 和 ProfileFixtures.xml CRLF 规范化提示来自既有工作区文件；本阶段未修改两者。

## 已知问题

本合同验证 XLSX 存储的公式，不调用公式计算引擎。最后的 Footer 若跨中间小计，仍需快照聚合或手写明确公式。

## 风险与回归关注点

第三方 Provider 应按 XLSX 标准写入公式节点及工作簿关系。合同只声明其实际检查的 XLSX 行为；远端 CI 尚未在本机执行。

## Reviewer 注意事项

合同包有一个新增公开入口，不属于八个生产程序集的 API 快照范围。新入口没有修改既有六项合同的调用签名。

## Git 状态

保留此前未提交的阶段改动；本阶段未执行 git add、commit、push、创建 PR、发布或修改版本号。

IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
RELEASE_STATUS: NOT_EVALUATED
OPEN_ACTIONABLE: 0
