# Continuation 15：独立 Entity Provider 合同测试包

## 范围

上一阶段的合同以源码链接交付。本阶段将同一六项 XLSX Entity Layout 合同打包为可引用的测试程序集，让第三方 Provider 测试工程只通过 NuGet 包接入。合同语义、Provider 运行代码和正式生产包版本保持不变；只在本地和 CI 构建包，不发布。

## 实施顺序

1. 建立双 TFM 的 `Bing.Offices.EntityProviderContracts` 测试包工程，编译现有合同源码，包依赖仅包含公开 Abstractions。沿用仓库现有 `2.0.0` 版本属性，不修改版本号。
2. 第三方 Provider 消费者移除源码链接，改为引用生成的合同包；执行全部六项合同，其中命名布局冲突验证导出目标保护。
3. 在 CI 包验证中独立打包合同，专用源供第三方消费者还原；原八个生产包计数保持不变。
4. 更新接入文档、生产符号追溯和执行记录。源码链接方式可作为旧用法保留，但推荐包引用。

## Change Impact 与验证

- ChangedProjects：新增测试包工程、第三方包消费者；CI 和文档。无 Provider、Abstractions 或其他运行时代码改动。
- ChangedPublicContracts：新增独立测试包的 `EntityLayoutProviderContractSuite`；既有生产程序集 API 不变。
- RiskLevel：MEDIUM；重点验证包内容、依赖和消费者确实从包加载合同，而非隐含源码引用。
- L0：合同包工程双 TFM 构建和本地打包；检查 nuspec、net6.0/net8.0 DLL 和依赖。
- L1：独立消费者从本地产物还原、构建并在 net6.0/net8.0 执行六项合同。
- L2：现有 Provider Contract 定向测试确认源码合同继续可用；Docs 测试及 CI 配置检查。
- L3：Release 构建只覆盖受影响工程；复用上一阶段未变的 Provider 常规测试证据，不重跑资源矩阵。

## 后续

Footer 公式/签字区等会改变 Provider 行为，另立下一阶段并安排职责级测试、双格式矩阵及 API 审批；本阶段不引入公式 DSL。旧版迁移分析器、Source Generator 和 DOM/Streaming 策略继续后置。
