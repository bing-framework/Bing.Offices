<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-ENTITY-PAGE-HEADER-CONTRACTS-20261004-001

# 实施执行报告

## 结果

完成 DESIGN-005：`Bing.Offices.EntityProviderContracts` 增加 `EntityLayoutProviderContractSuite.VerifyPageHeaders`，以公开 Abstractions 验证重复页眉跨 Provider 行为。官方 NPOI、ClosedXML 共同运行；第三方包消费者通过 NuGet 调用同一入口。本阶段没有修改生产逻辑或生产公共 API。

## 范围与验证

- 共享合同覆盖空、单、多页往返、页码、复制页表头、占用冲突时目标流保护，以及 marker 不匹配时的工作表原始坐标。
- 官方 Provider 合同测试使用公开 options 重载验证重复页眉工作簿的输入字节限制；NPOI 在 `Open` 阶段、ClosedXML 在 `Preflight` 阶段报资源错误。
- NPOI XLS/HSSF、XLSX/XSSF 与 ClosedXML XLSX 的引擎结构、样式、合并、边界测试保留在各自项目，未以共享 XLSX 合同替代。
- 新增职责用例在 net6.0、net8.0 各 4 项通过；合同 NuGet 包打包通过，第三方消费者在两个 TFM 均输出 `third-party-public-only-provider-ok`。
- `Bing.Offices.sln` Release `Category!=Large` 全量常规测试通过。仅出现已有 `Microsoft.Bcl.Memory` 对 net6.0 的支持提示。
- `Bing.Offices.sln` Release 构建和独立目录生产包打包通过；使用当前 DLL、独立生产包、现有成员审批与基线执行 API Snapshot，比对 net6.0、net8.0 均通过，`api-diff.json` 无成员差异。
- `git diff --check` 通过。本轮涉及的 C# 文件均为严格 UTF-8 BOM + LF，项目文件为无 BOM + CRLF，其余文本为无 BOM + LF；均有单个末尾换行且无混合行尾。全局 `git diff --stat` 包含本任务开始前的其他未提交工作。

## API 与兼容性

生产程序集本轮无成员变化。测试合同包在 net6.0、net8.0 新增公开静态方法 `EntityLayoutProviderContractSuite.VerifyPageHeaders(IExcelEntityExporter, IExcelEntityImporter)`；原有入口及调用方式不变。使用该能力的第三方 Provider 应在声明支持 Entity 实际重复页眉时运行此合同。

## 文件与交付

新增独立合同实现文件、合同测试和任务追溯；更新合同包项目文件、NuGet 消费者与 `docs/excel/entity-provider-contracts.md`。不提交、推送、发布或修改版本号。
