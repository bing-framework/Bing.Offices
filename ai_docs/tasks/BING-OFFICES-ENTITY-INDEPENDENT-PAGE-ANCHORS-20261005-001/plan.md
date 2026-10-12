# Entity 独立逐页页眉与签字区命名锚点

Task ID：`BING-OFFICES-ENTITY-INDEPENDENT-PAGE-ANCHORS-20261005-001`

## 目标与真实现状

执行后续 F-001。当前 PageSubtotalNamed 只对存在下一页的中间小计 marker 生成名称，FooterNamed 只定位最终 Footer；PageHeader 支持实际复制，但没有逐页名称。首页、末页、无小计和空明细不能由现有分页名称系列完整定位。

已确认入口为 Abstractions 的 ExcelEntityListRegionBuilder、ExcelEntityListRegion、ExcelEntityLayoutBuilder；物理布局由 src/provider-shared/ExcelEntityRowPlan 及现有 Provider 分页分支处理，导入跳过由 ExcelEntitySubtotalImportScanner 处理。NPOI 和 ClosedXML 均已支持重复页眉、分组分页和中间页小计，并有完整成功、边界与资源测试。不能仅增加名称而继续让签字区依赖中间小计、缺少末页内容。

本阶段为 NPOI XLS/HSSF、XLSX/XSSF 和 ClosedXML XLSX 提供独立逐页页眉命名，以及每个逻辑页都有的 Footer/签字区和命名。复用现有 Footer 单元格、聚合、样式与合并描述，不建立第二套单据布局。所有旧公开签名保持不变，不公开内部辅助类型、不增加生产友元。

## 公共契约与兼容性

在 ExcelEntityListRegionBuilder<TItem> 新增：

```csharp
PageHeaderNamed(string anchorPrefix, string sourceRange,
    string markerAddress, string markerText,
    Action<ExcelEntityPageHeaderBuilder<TItem>> configure = null,
    ExcelEntityPageAnchorStaleNamePolicy staleNames = ExcelEntityPageAnchorStaleNamePolicy.Reject)

PageFooter(string markerText,
    Action<ExcelEntityListFooterBuilder<TItem>> configure = null)

PageFooterNamed(string anchorPrefix, string markerText,
    Action<ExcelEntityListFooterBuilder<TItem>> configure = null,
    ExcelEntityPageAnchorStaleNamePolicy staleNames = ExcelEntityPageAnchorStaleNamePolicy.Reject)
```

新增 ExcelEntityPageAnchorStaleNamePolicy 枚举：Reject、RemoveObsolete。单独文件归入 Entities/Anchors，公开命名空间仍为 Bing.Offices.Entities。

ExcelEntityListRegion<TEntity> 增加五个只读、EditorBrowsable(Never) Provider SPI 属性：PageHeaderAnchorPrefix、PageHeaderAnchorStaleNamePolicy、PageFooter、PageFooterAnchorPrefix、PageFooterAnchorStaleNamePolicy。描述经布局构建快照、命名起点解析和内部构造链完整传递。没有名称的 PageFooter 只负责内容，不创建名称；PageHeader 原入口保持原行为。

新增入口先完整验证参数，再改变构建器状态；重复配置、无效策略或配置委托失败不留下半成品。旧 PageSubtotalNamed 的名称及过期名称拒绝语义保持原样。

### 逻辑页与每页 Footer

- 两个每页入口要求 PageBreak。rowsPerPage 只统计明细；有数据时按容量划分，空明细明确视为一个逻辑页。没有 PageHeader 时也可以声明 PageFooterNamed，不要求 PageSubtotal。
- PageHeaderNamed 为已有 PageHeader 的 marker 命名，包含首页、末页、单页及空明细页；不在没有实际页眉时制造页眉名称，也不改变与 PrintTitleRows 的互斥规则。
- PageFooter 在每页最后一个明细及该边界分组小计、分页小计之后写入，在分页符和下一页页眉之前；末页也写入，随后再写最终 Footer。组未结束时不提前生成组小计。
- PageFooter 的相对 A1 为其 marker，GapRows 使用原有 Footer 语义。固定值、按当前页明细快照求值、样式、数字格式及合并全部复用现有 Footer 构建器。可直接声明签字、审核、日期等内容。跨小计明细求和公式仍只允许最终 Footer；不悄悄扩大公式 DSL。
- 空明细页仍写首页页眉及一个每页 Footer，聚合接收空只读快照；没有配置的区域不生成。最后一页恰好满容量也只生成一个每页 Footer。
- marker 与其它页眉、小计、最终 Footer 标记不得歧义；导入仍依靠标记和页容量，名称服务外部定位，不替代内容校验。

### 名称及模板语义

- 名称为 `{prefix}_{pageIndex:000}`，从 001 开始，超过 999 按自然位数扩展。分别指向实际页眉 marker 与每页 Footer marker，不指向逻辑行号或可变明细。
- 前缀及全部生成名称遵守现有名称规则与 255 字符边界；大小写不敏感校验系列之间、静态单格锚点、最终 Footer 及分页小计系列的冲突。
- 模板中目标 Sheet 的同名单格名称允许重定位并保留作用域。跨 Sheet、失效引用、公式、外部引用、非单格以及同名多作用域歧义拒绝；无关名称不改动。
- Reject 默认对本次未使用的旧系列名称报配置错误，避免静默留下过期页引用。RemoveObsolete 显式选择时，只删除属于当前系列、合法且指向目标 Sheet 的旧单格名称；任何歧义或非法旧引用仍拒绝。清理发生在输出副本内部，不修改输入模板流或磁盘模板。
- 先完成所有系列、作用域和引用预检，再重定位或清理；失败时旧输出、请求快照和输入模板保持原有合同。

## P1 / ANCHOR-001：模型、构建与快照

目标：增加最小公开扩展，保留原 API 并固化名称及状态合同。

已确认修改文件：ExcelEntityListRegionBuilder.cs、ExcelEntityListRegion.cs、ExcelEntityLayoutBuilder.cs；候选新增文件为 Entities/Anchors/ExcelEntityPageAnchorStaleNamePolicy.cs。内部地址和名称辅助复用现有能力命名空间，不复制 Provider 源码绕过边界。

实施步骤：实现三个入口、枚举和五个 SPI 属性；校验先于状态提交；传递快照及 WithResolvedStart；检查 PageBreak、保留 marker、系列与静态名称冲突；旧构造、入口及原默认值均保留。

验证：职责测试覆盖有效/无效前缀、未知枚举、重复配置、异常委托、构建失败后状态、快照隔离、重用渲染及命名起点。明确断言实际异常类型，不只断言“抛异常”。

依赖：CLIENT 正式收口；风险：新增描述未传过命名解析或错误调用污染构建器。验收：新描述可由公开 API 读取，不引入友元，旧调用继续编译，受影响成员中文 XML 完整。

## P2 / ANCHOR-002：共享物理行与导入识别

目标：所有页的签字区、页眉、小计和名称使用同一物理规划。

已确认文件：src/provider-shared/ExcelEntityRowPlan.cs、ExcelEntitySubtotalImportScanner.cs、ExcelEntityPageAnchorNames.cs；必要时单独增加内部单类型描述文件，保持一个文件只承载对应职责。

实施步骤：增加每页 Footer 行段及页面归属；在页末按组小计→中间分页小计→每页 Footer→分页符→下页页眉的顺序规划，末页每页 Footer 之后才是最终 Footer。用 long 做累加，保留逻辑容量，检查取消；空明细生成一个可识别签字区。带新 PageFooter 的 Provider 路径统一进入共享规划，不重写无关旧快速分支。

扩展扫描器按 marker、占用高度、GapRows 和明细容量验证每页 Footer，包含最后一页及空页。缺失、重复、错位、额外区域与非法高度拒绝，缺失位置根据物理规划给出原始 Sheet/一基行列。分组跨页、组页同边界和无表头均遵守同一顺序。NPOI 临时投影和 ClosedXML 直接读取跳过整个区域，保持原生校验→转换→配置校验及 Unique 回滚。

验证：完整行段/导入职责覆盖 0/1/整页/多页、无小计、有小计、跨组、GapRows、末页满容量、命名起点漂移、损坏 marker、保留 End、原始错误坐标及取消。

依赖：ANCHOR-001；风险：重复 marker 被当明细、末页丢失或页码偏移。验收：完整工作簿往返不包含签字区假明细，已有分页小计名称位置不漂移。

## P3 / ANCHOR-003：双 Provider 名称与模板

目标：三种格式的真实内容、名称、模板、资源合同一致。

已确认文件：NpoiEntityExportExecutor.cs、NpoiEntityImportExecutor.cs、NpoiEntityLayoutSupport.cs，ClosedXmlEntityLayoutExecutor.cs。沿用 PreparePageAnchorNames/UpdateFooterAnchor 及既有名称解析能力；如抽取辅助类型，放在对应 Entities 能力目录，类型保持 internal。

实施步骤：根据共享规划形成首页及后续页眉、每页 Footer 的单格名称；全部系列预检后写入内容与名称。实现显式旧名称清理，保留同名单格作用域和无关模板设置。预检所有实际行列、动态/计算列宽度、合并、固定字段及同 Sheet 其它区域的运行时占用；格式溢出和名称错误在目标流写入前失败。异步外围 IO、DOM 准入、默认提交、异常观察、暂存和调用方流所有权不变。

验证：NPOI XLS/HSSF、XLSX/XSSF 均有真实成功与失败；ClosedXML XLSX 同等覆盖。断言全部明细、每页签字值/合并/样式、名称数量与每个实际引用、页眉页码、小计、分页符、最终 Footer 及输入模板不变。两种旧名称策略分别覆盖缩短、局部名称、跨 Sheet、歧义、失效、公式及范围；清理失败后旧输出完整。

依赖：ANCHOR-002；风险：局部名称覆盖全局、清理用户无关名称或把格式限制推迟到序列化。验收：三个格式完整往返与结构化失败通过，同步/异步结果结构一致。

## P4 / ANCHOR-004：业务合同、资源与包消费者

目标：用真实采购单签字区证明独立能力，不仅检查 API 存在。

候选新测试文件：两个 Provider 测试目录的 EntityIndependentPageAnchorTest.cs；模型职责可进入 ProviderContract.Tests 的独立文件。复用已有共享 Entity 合同入口，并在第三方/双 TFM NuGet 消费者演示新描述与两种名称系列，无生产项目引用或友元。

矩阵：无 PageSubtotal、多页首页/末页、空明细、打印标题配签字 Footer、实际页眉配签字 Footer、分组分页同边界、动态/计算列、固定字段、多区域、模板样式和命名起点。导入校验错误保留原始坐标，所有错误及成功数据完整断言。

资源：聚合委托异常、预取消、布局/内容写入中取消、提交失败与最大行列边界；核对旧目标、暂存、调用方流及异常观察一次。只使用有界数据和现有资源限制，不重跑无关 Large 矩阵，不用 Task.Run/.Result/.Wait 伪异步，不引入无证据的缓存。

依赖：ANCHOR-003；风险：共享合同掩盖引擎差异。验收：每个 Provider 有独立职责；包消费者在 net6/net8 完整执行，旧 Fluent 与旧分页名称合同不回退。

## P5 / ANCHOR-005：API、文档与交付

目标：源码、公开契约、迁移及证据一致。

更新 dynamic-columns.md、Provider 矩阵、Entity 合同说明、NuGet 迁移示例，以及本任务 api-approval.md、双 TFM API 差异、symbol-test-map.md、execution.md。示例展示每页页眉与签字区名称，说明逻辑页与真实打印页不同、名称定位单格、不计算公式、显式清理只处理当前系列。

本轮成员审批依据用户已持续授权的 F-001 实施：只新增上述三个 Fluent 方法、五个 SPI 属性及枚举/枚举值；无删除、旧签名变化或新生产依赖。先生成并检查双 TFM 成员差异，再按既有流程更新快照和候选身份，不覆盖基线掩盖差异。

实际验证入口：dotnet test 两个 Provider/ProviderContract/Integration/Docs 项目，Release、net6.0/net8.0、--no-restore -m:1 -p:UseSharedCompilation=false；职责先用 FullyQualifiedName~EntityIndependentPageAnchor，再执行受影响回归。使用仓库 build/ApiSnapshot、现有 Consumer、第三方消费者和 .github/workflows/ci.yml 的常规 Category!=Large 门禁。Release 构建及串行 NuGet pack 后执行隔离包消费者，保存命令退出码和 TRX。

最后执行 git diff --check、差异统计、严格 UTF-8/BOM/EOL/EOF 检查并独立复核。C# BOM/LF、csproj/ps1 CRLF、文档与 JSON 无 BOM/LF；保留前序工作区和历史证据。使用 task-finish --no-notify 仅收口本地状态，不提交、推送、PR、发布、改版本或发送外部通知。

依赖：ANCHOR-004；风险：API 额外变化或旧消费者兼容回退。验收：仅审批内成员新增，全部相关门禁和独立复核通过，追溯到真实方法与产物。

## 范围与后续

不加入高级透视、NPOI 透视缓存、透视模板更新、前向失败工作簿、低内存 Entity 或公式计算引擎；它们仍按 F-002～007 推进。无关绘图复制、真实打印渲染和商业环境不由名称定位替代。计划完成不等于本阶段已实施；先写计划，再由执行阶段注册状态并实现。用户已授权按 TODO 顺序继续，常规实现决策无需重复请求确认。
