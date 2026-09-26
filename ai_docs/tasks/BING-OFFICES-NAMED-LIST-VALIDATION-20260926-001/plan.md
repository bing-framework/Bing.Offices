# 开源校验后续能力与证据补齐

Task ID：`BING-OFFICES-NAMED-LIST-VALIDATION-20260926-001`

## 范围与依据

用户要求继续完善后续能力，延续开源优先、暂不处理 Aspose。上一任务 F01–F06 保持已完成，本任务不重写其历史证据。

当前 ClosedXmlWorkbookValidationPipeline.ResolveListValues 只接受显式文本和直接 A1 区域，存在有效名称时仍 Unsupported。已有 Report/Fail、同步/异步入口及配置校验顺序可复用。补齐名称范围列表导入校验作为高级校验第一步，不宣称完整公式计算。

不新增公共 API、依赖或版本，不修改 NPOI 行为、导出原生校验模型、商业能力和公式处理器。既有未知名称仍 Unsupported。

## N01：有界名称列表

- 目标：ClosedXML 原生列表校验解析 Workbook/当前 Sheet 名称，局部名称优先，支持引用其他 Sheet 的单行/单列绝对有界区域及单个单元格。
- 确认文件：`src/Bing.Offices.ClosedXml/Bing/Offices/Imports/ClosedXmlWorkbookValidationPipeline.cs`；候选新增私有解析 helper 和独立测试类。
- 步骤：分离直接区域与名称解析；使用引擎的名称集合读取定义，不执行名称公式；解析一个明确区域；逐单元格比较并检查取消，避免为每行复制整个列表。
- 拒绝：未知名称、失效/外部引用、动态公式、别名链/循环、联合区域、整行整列、二维名称区域、相对名称引用。无效局部名称不得回退到同名全局名称。
- 兼容：不删除或改变公开成员，直接区域原有语义保持；复杂 Custom 和公式重新计算继续 Unsupported。
- 风险：Sheet 引号转义、名称优先级、公式触发计算、扫描资源。名称只解析绝对有界地址；不访问网络，不求值动态名称，不建立全局缓存。
- 验证：独立职责测试覆盖全局/局部/遮蔽、跨 Sheet/含空格及单引号、合法/非法/空值、Report/Fail、同步/异步、拒绝集合、取消和调用方流所有权；通过公共导入入口保存后重新打开测试。
- 验收：有效名称按真实允许值判断；非法实体及 Unsupported 分别产生既有结构化结果；没有静默回退。

## N02：可复核的运行环境记录

- 目标：修正上一报告把串行脚本配置写成实测并发计数的问题；以后 Docker 运行自动保存 Engine 版本与镜像身份。
- 确认文件：`scripts/verify-docker.sh`、上一任务 `resource-report.md`；新增本任务执行记录。
- 步骤：报告明确“脚本串行、无后台队列”为实现配置而非测量字段；脚本保存 docker version JSON、image inspect 精简身份；不向旧原始日志补造历史字段。
- 验证：shell 语法检查、真实 metadata 命令及一个受影响 Linux 校验测试运行；不重复不受影响的 9 组容量。
- 风险：Windows Git Bash 路径转换，继续复用现有策略。
- 验收：未来运行产物有可读版本/镜像身份，旧报告测量与推断清晰分开。

## N03：验证、文档与收口

- 先定向测试，再 ClosedXML 双 TFM 全项目；ProviderContract 受影响校验合同与 Linux 定向验证。
- 不改公共 API 形状，运行既有 PublicApiContractTest 双 TFM；无需批准新成员或覆盖历史 API 制品身份。
- 生产导出热路径未变，上一容量记录复用，不重复全 Solution Large/百万行探针。
- 更新 Provider 校验边界、生产符号到测试方法映射、execution.md；独立 Reviewer 只审本任务差异；仅 OPEN_ACTIONABLE>0 修复。
- 检查 git diff --check、严格 UTF-8/BOM/EOL；.cs 保持 BOM+LF，其余按仓库契约。
- 不自动 commit/push/PR/发布，不发送外部通知。

## 完成标准

N01 的公共入口与职责测试通过；N02 记录可复核；N03 双 TFM 与 Linux 定向证据齐全。后续公式计算、导出名称列表 API、商业渲染、透视表仍另阶段处理，不合并为本轮承诺。
