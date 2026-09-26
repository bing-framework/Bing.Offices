<!-- AI_REVIEW_STATUS: PASS -->
AI_TASK_ID: BING-OFFICES-NAMED-LIST-VALIDATION-20260926-001
AI_REVIEWED_AT: 2026-09-26T15:30:41.5916876Z

# 独立审查

## 审查范围

本次只审查本任务的 Named List 校验实现、直接测试和 Docker 元数据脚本：

- `src/Bing.Offices.ClosedXml/Bing/Offices/Imports/ClosedXmlWorkbookValidationPipeline.cs`
- `tests/Bing.Offices.ClosedXml.Tests/NamedListValidationTest.cs`
- `scripts/verify-docker.sh`
- 本任务的 `plan.md`、`execution.md`、`validation-support-matrix.md` 和 `production-symbol-test-map.md`

未修改生产代码、测试代码或脚本，也未重复审查之前任务的其他 Provider 实现。

## 影响分析

| 项目 | 结论 |
| --- | --- |
| 变更运行路径 | ClosedXML WorkbookValidation 的 List 规则解析、定义名称解析和范围匹配 |
| 受影响调用方 | ClosedXML 配置导入器、ProviderContract 校验合同、Docker smoke 测试 |
| 公共 API | 无新增、删除或签名变更 |
| TFM | net6.0、net8.0 |
| 资源/取消 | 命名范围逐单元检查取消；输入流所有权由直接合同覆盖 |
| 风险等级 | 中等，集中在 Workbook 规则边界解析 |

## 验收矩阵

| 计划项 | 结果 | 证据 |
| --- | --- | --- |
| N01 Named List 生产实现 | PASS | 全局/局部名称、局部遮蔽、单元格/单行/单列有界绝对范围、转义 Sheet 名称、非法/动态/别名/循环/联合/二维/整行整列/相对引用拒绝均有实现和直接测试 |
| N01 合同语义 | PASS | `targeted-final2-net6/named-list-net6.trx` 22/22；`targeted-final2-net8/named-list-net8.trx` 22/22；同步/异步、Report/Fail、空值、取消和输入所有权均被覆盖 |
| N02 Docker 元数据 | PASS | `scripts/verify-docker.sh --metadata` 写出 Docker 版本和镜像身份；Git Bash `bash.exe -n` 语法检查退出码 0；元数据 JSON 可解析 |
| N03 分层验证 | PASS | Windows ClosedXML net6/net8 各 137/137；Linux Docker ClosedXML net6/net8 各 137/137；ProviderContract 各 18/18；Public API 各 9/9 |

## 实现核对

`TryResolveNamedRange` 先解析工作表局部定义；存在局部定义时不回退到同名全局定义，符合局部名称遮蔽和“局部无效即失败”的要求。`TryResolveDefinedName` 仅接受绝对的单元格、单行或单列有界区域，并通过 `TrySplitReference` 拒绝外部引用、联合区域和未闭合的 Sheet 引号。命名范围的每个已使用单元格都会检查公式和取消状态，避免先命中值后绕过后续不支持内容。直接 A1 范围和显式列表仍保留其既有语义。

公共导入器路径将无效值记录为 `WorkbookValidation` 错误，并按 Report/Fail 策略处理不支持规则；测试断言了错误码、Sheet、行、列和属性，而不是只断言结果数量。

脚本的 `--metadata` 路径只采集当前主镜像的 Docker 版本和镜像身份，不伪造历史容量或并发数据。主 smoke/fontless 路径继续使用非 root、只读输入、临时目录和内存限制。

## Findings

`OPEN_ACTIONABLE: 0`

未发现需要修复的生产缺陷、合同回归、公共 API 破坏或 Docker 脚本错误。因此不创建 `FIX-*` 项，也不进入修复轮次。

## 已确认边界（非缺陷）

1. ClosedXML 解析器无法构造外部定义名称。实现和矩阵将该情况明确列为引擎解析边界，未错误宣称可在 Report/Fail 策略下处理；这不是静默回退或误判。
2. 取消检查位于命名范围逐单元扫描循环中，源码映射和预取消合同已验证；执行报告明确记录没有独立的“循环中途”测试。该项是证据粒度说明，不构成当前生产行为缺陷。
3. 早期 `targeted-net8` 历史结果不作为最终结论；最终 `targeted-final2` 双 TFM 均为 22/22，且与执行报告一致。

## 验证记录

- `C:\Program Files\Git\usr\bin\bash.exe -n scripts/verify-docker.sh`：退出码 0。
- Docker 版本元数据解析成功：Docker Engine 29.4.0，Linux amd64。
- 主镜像身份元数据解析成功：`sha256:ebed27f7de289a3d5c2f77ab0b19a448458db074e8192aada1a566748eb1dbad`。
- 相关源码、测试、脚本和任务文档通过严格 UTF-8、BOM、行尾检查；`git diff --check` 无输出。

## 结论

本任务的 Named List 解析、Unsupported 边界、Report/Fail 语义、同步/异步合同和 Docker 元数据采集均达到计划要求。实现、测试、文档和交付证据均为 `PASS`，当前没有 `OPEN_ACTIONABLE`，审查结束。
