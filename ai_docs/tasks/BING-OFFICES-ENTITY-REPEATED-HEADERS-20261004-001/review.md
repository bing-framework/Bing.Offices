<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BING-OFFICES-ENTITY-REPEATED-HEADERS-20261004-001
AI_REVIEWED_AT: 2026-10-04T10:37:42Z

# 重复页眉计划复核与 TODO

## 结论

按用户批准的完整计划、仓库 plan.md、execution.md、当前源码/差异及包行为核对：打印标题和实际页眉两个主能力均已接入，不存在整项入口缺失；仍有 **4 项已复现缺陷和 1 项组合验收补齐任务**，不能认定计划完全验收。

- REVIEW_STATUS：NEEDS_FIX；IMPLEMENTATION_STATUS：PARTIAL；TEST_STATUS：PARTIAL。
- OPEN_ACTIONABLE：5，其中 MUST_FIX 2、SHOULD_FIX 3；外部阻塞 0。
- 本轮只审查，未修改生产源码、正式测试、plan.md 或 execution.md。
- 前轮审查保存在 review-round1.md。原 FIX-001～FIX-007 有修复及回归证据；占用检查、错误定位、验收覆盖仍有以下遗漏组合。本轮使用新的 FIX 编号，避免覆盖历史问题含义。

## 计划验收矩阵

| 部分 | 状态 | 结论 |
| --- | --- | --- |
| PrintTitleRows | PASS | 两个 Provider 已写打印元数据，无需 PageBreak、不增加明细行；已有职责测试和消费者 |
| PageHeader API、上下文与快照 | PASS | 构建器、页码/总页数/只读集合、相对 Cell/Merge、隐藏 SPI 已接入 |
| 复制与共享行规划 | PARTIAL | 三种格式正常往返、分页小计/分组及页锚点已有证据；无模板目标占用漏检见 FIX-008 |
| 导入识别与坐标 | PARTIAL | 首页/重复 marker、普通并列表及无 End 并列表已修复；剩余 FIX-009～FIX-011 |
| 验证、API、包和文档 | PARTIAL | 前轮常规门禁及包验证通过；组合/资源证据仍缺 FIX-012 |

## 证据与验证边界

复用最后一轮执行结果：重复页眉职责测试 NPOI 双 TFM 各 32 项、ClosedXML 各 21 项；Provider 常规测试分别各 891/306 项；解决方案 Category!=Large、串行 pack、双 TFM Consumer、API snapshot 比较均通过。本轮没有重复无变化的全量测试。

另运行 `artifacts/review-round2/Probe.csproj`：隔离 net8.0 程序引用上一轮最终本地 NuGet 包和消费者缓存，无生产 ProjectReference。先验证正常三行往返，再运行以下场景：

| 场景 | NPOI XLS | NPOI XLSX | ClosedXML XLSX |
| --- | --- | --- | --- |
| 无 End、无小计的普通页眉 | 成功，3 行 | 成功，3 行 | 成功，3 行 |
| End(C20)，实际最后明细在第 7 行 | ConfigurationInvalid，S!A21 | 同左 | 同左 |
| 无分页小计，第二页 A5 marker 改为 BROKEN | 错误坐标 S!A8 | 同左 | 同左 |
| 无表头命名起点 B3 改为 B4，原数据留在 B3 | IsSuccess=true，0 行 | 同左 | 同左 |
| 无模板固定 A5=BUSINESS，与第二页 marker 重叠 | 本轮未运行 | 输出前拒绝 | 导出成功，A5=HEAD |

新探针仅运行 net8.0；修复阶段仍须加入正式双 TFM 回归。暂存区无差异，工作区包含多项历史任务改动，未全部归于当前任务。

## 计划内 TODO

### FIX-008：ClosedXML 无模板页眉覆盖固定字段

- [ ] P0；严重程度 HIGH；处理要求 MUST_FIX；状态 OPEN_ACTIONABLE。
- 问题：固定字段先写入工作簿，复制页眉的目标内容检查却只在模板模式执行。
- 证据：`ClosedXmlEntityLayoutExecutor.cs:774` 的 `if (!template) continue`；固定 A5、列表 B2、源 A1:C2、每页 2 条、3 条明细时，A5 从 BUSINESS 变为 HEAD。NPOI 同例拒绝输出。
- 影响：静默覆盖业务字段，Provider 行为不一致。
- 修复目标/要求：对所有复制页眉目标检查已配置固定字段、内容和合并占用，覆盖普通/命名字段及空值字段；不能只凭非空单元格识别占用。
- 验证方式：双 Provider、NPOI 两种格式覆盖冲突失败和合法不重叠，核对输出前失败、原目标和流所有权。

### FIX-009：命名起点改变后静默漏读

- [ ] P0；严重程度 HIGH；处理要求 MUST_FIX；状态 OPEN_ACTIONABLE。
- 问题：命名起点在构建时推迟邻接校验，导出补验，但导入解析后未重新验证源末行紧邻实际明细。
- 证据：`NpoiEntityImportExecutor.cs:173`、`ClosedXmlEntityLayoutExecutor.cs:250`。源 A1:C2、Header(false)、命名起点 B3，导出一条数据后仅将名称改为 B4，三个组合均成功导入 0 行。
- 影响：工作簿名称改变可令真实明细被忽略，违反“命名列表起点解析后仍须满足相邻关系”。
- 修复目标/要求：在导入命名解析后验证源范围、表头模式与实际起点的邻接关系，返回原始位置的配置错误，不能依赖表头匹配间接拦截。
- 验证方式：合法命名往返及起点前移/后移，覆盖有/无表头、空/单/多页、XLS/XLSX。

### FIX-010：End 预留空行被算作明细

- [ ] P1；严重程度 MEDIUM；处理要求 SHOULD_FIX；状态 OPEN_ACTIONABLE。
- 问题：无 Footer 的页眉列表直接用 End.Row 作为扫描末行，末尾空行也被计入分页容量。
- 证据：`NpoiEntityImportExecutor.cs:503`、`ClosedXmlEntityLayoutExecutor.cs:993` 优先采用 End；`ExcelEntitySubtotalImportScanner.cs:120` 累计末尾行差。End(C20)、实际最后明细第 7 行，均报 S!A21 页眉位置无效。
- 影响：常见预留填写区导出后无法按原布局往返。上一轮无 End 修复未覆盖此组合。
- 修复目标/要求：区分声明上界与实际内容终点，在 End 内确定真实终点，同时保留缺失页眉、重复 marker、越界的拒绝语义，不能通过截断掩盖损坏内容。
- 验证方式：空/单/多页、End 恰好到末行/预留容量、尾部只有样式、相邻更长列表、真实缺失 marker。

### FIX-011：无分页小计时缺失 marker 定位不准

- [ ] P1；严重程度 MEDIUM；处理要求 SHOULD_FIX；状态 OPEN_ACTIONABLE。
- 问题：没有 PageSubtotal 时，缺失页眉的最终错误使用 endRowExclusive 定位，而非本应出现的 marker。
- 证据：`ExcelEntitySubtotalImportScanner.cs:120-127`；第二页 A5 改为 BROKEN，返回 S!A8。已有精确坐标测试带分页小计，走 awaitingHeader 分支，未覆盖此处。
- 影响：异常被识别，但应用无法准确标记损坏页眉。
- 修复目标/要求：按页容量、前序行段及 marker 相对偏移计算缺失位置；有/无小计保持同一原始坐标契约。
- 验证方式：无小计、有分组、多页、marker 不在源首行，精确断言 Code、Stage、Sheet、Row、Column。

### FIX-012：补齐组合与资源验收

- [ ] P1；严重程度 MEDIUM；处理要求 SHOULD_FIX；状态 OPEN_ACTIONABLE。
- 问题：现有页眉测试主要使用固定列、绝对起点和同步成功路径。独立功能测试不等于页眉组合已验收。
- 证据：两个 `EntityRepeatedPageHeaderTest` 与 Consumer 的页眉场景未覆盖 ListRegionNamed、Header(false)、动态/计算列组合；异步页眉用例主要是取消/预检失败。最大行溢出和 marker 移位不能替代完整非法高度矩阵。通用提交器/失败工作簿测试也不能证明页眉专属包装路径的观察次数。
- 影响：全套测试通过仍遗漏上述已复现组合故障。
- 修复目标/要求：补命名起点、无表头、动态/计算列最终物理范围、非法高度、异步成功往返；补页眉导出的真实提交失败、原文件完整、临时文件清理、异常观察一次。
- 验证方式：先职责级双 TFM/三格式矩阵，再受影响合同、Consumer、API/Release/常规门禁；同步更新成员追溯和 execution，不重复无关性能矩阵。

## 欠缺考虑的设计 TODO

以下为建议项，不混入上述 5 项验收欠账：

- [ ] DESIGN-001 / P1：明确 marker 与业务值同值的处理。当前按标记列文本匹配，若它同时是数据列，业务值等于 marker 会产生歧义。定义保留列/保留值、导出前冲突检查或位置优先识别规则。
- [ ] DESIGN-002 / P2：明确普通 Fluent Cell/Merge 覆盖列表表头的边界。marker 已受保护，普通覆盖和合并仍需要定义哪些标题定制合法、哪些会破坏同布局导入。
- [ ] DESIGN-003 / P2：细化复制清单：富文本、超链接、批注、校验、条件格式、隐藏行、锁定状态分别承诺保留、不复制还是拒绝。普通值、样式、行高、公式文本、完整合并已有约定；绘图对象仍明确不复制。
- [ ] DESIGN-004 / P2：定向验证大跨度扫描和取消响应。末行回溯与共享 marker 扫描没有 CancellationToken 参数；先测稀疏大表、大 End 和资源限额生效顺序，再决定按块取消或按已用单元格扫描。本轮没有测得实际性能回归。
- [ ] DESIGN-005 / P2：将页眉行为矩阵提取成可复用 Provider 合同测试，统一占用、坐标及资源失败场景，同时保留引擎特有的结构断言，降低两份实现产生行为差异的风险。
- [ ] DESIGN-006 / P3：公共透视表模型、公式计算与真实打印渲染验收另行设计。它们不是本阶段未实现项；逻辑页数不能直接承诺为纸张/缩放/行高作用后的打印页数。

## 推进顺序

先 FIX-008、FIX-009 消除覆盖和漏读，再 FIX-010、FIX-011 修正边界与定位；FIX-012 随修复补齐，最后门禁和复审。设计建议优先 DESIGN-001/002，再考虑复制元数据、扫描取消和合同测试；透视表/渲染独立规划。

本轮未修改业务逻辑，未 commit、push 或发布。
