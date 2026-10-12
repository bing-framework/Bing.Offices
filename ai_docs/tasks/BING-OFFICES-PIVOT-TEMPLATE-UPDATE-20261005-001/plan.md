# 已有透视模板保留与更新

Task ID：`BING-OFFICES-PIVOT-TEMPLATE-UPDATE-20261005-001`

## 范围与授权

沿用用户“继续实现 TODO，按建议顺序推进”的 F-004。本阶段目标是 NPOI XLSX、ClosedXML XLSX 的已有原生透视模板保留与显式更新，不用普通模板创建新透视代替。F-002/003 的客户端 Count 失败、最终独立验收以及 F-005～007、E-001/E-002 保留，不缩小完整 Goal。先落实本计划，按阶段实施，不自动提交、推送、发布或修改版本。

## 当前真实调用链

- WorkbookBuilder.UseTemplate 保存输入及所有权；Request 是不可变执行描述；现有 PivotTable 只创建新目标 Sheet。
- NpoiReportWriter.ApplyPivotTables 在有新透视时根据 xl/pivotTables/、xl/pivotCache/ 标准路径拒绝已有部件；来源预检、类型化完整缓存、原生字段和一次 ZIP 转写已接入。
- ClosedXmlTemplatePreflight.Validate 根据 ZIP 标准路径拒绝 Pivot 部件，后续 XLWorkbook 加载和 SaveAs 是完整 DOM 序列化。
- 两个 ReportWriter 仍拒绝已存在 TargetSheet。现有业务写入覆盖来源数据，不代表旧透视缓存、字段索引和未知 XML 配置可无损更新。
- 当前包已有基础/高级完整缓存，两个 Provider 的基础/高级新建模型可复用；Count 客户端数值失败不作为模板能力成功证据。

## 公共入口与兼容决策

复用 ExcelPivotTableDefinition，不新建第二套字段/聚合描述。

候选新增成员（实施前建立成员级审批）：

1. ExcelWorkbookExportBuilder.PreserveTemplatePivotTables()：显式保留模板内已有透视；不改变其定义或自动刷新缓存。
2. ExcelWorkbookExportBuilder.UpdatePivotTable(ExcelPivotTableDefinition definition)：按名称和 TargetSheet 精确定位已有透视并更新来源、轴、度量及筛选，隐含开启保留其它透视。
3. Request 隐藏于 IntelliSense 的只读 PreserveTemplatePivotTables 与 PivotTableUpdates SPI；深快照规则沿用现有 Clone。
4. 新增细粒度 TemplatePivotTables 能力位仅在主链及验收成立后声明，格式限定 XLSX。

默认旧入口行为、旧签名、旧命名空间保持不变。新建和更新名称在请求内不得冲突；更新目标须在模板中存在且唯一。模板存在重名透视、未知目标或来源缺失时返回配置错误，不静默新建。源须为有界字面值，公式/Excel 错误拒绝规则不变。

仅显式保留/更新模式允许没有 AddSheet 的模板请求；普通 Workbook 至少一个 Sheet 的旧规则保留。不存在模板时这些入口在 Build 阶段失败。保留模式不承诺缓存与之后修改的业务来源同步，文档明确通过更新入口刷新；现有缓存不会被当成新结果。

## P1 / PTPL-001：OPC 清单与依赖图

目标：在任何 Provider 重写前确定实际透视、缓存、来源 Sheet 和依赖关系，不依赖部件编号或标准目录。

已确认修改范围：两个 Provider 的模板预检入口、项目编译链接。候选新增 provider-shared 的单职责内部类型：ExcelPivotTemplateParts（发现）、ExcelPivotTemplateManifest（关系清单）、ExcelPivotTemplateEntry/Cache（不可变描述）。公开生产程序集之间不使用友元或反射。

步骤：从 ContentTypes 和关系类型发现 PivotTable/CacheDefinition/CacheRecords；解析 workbook cacheId、Sheet 所属及实际相对 URI，验证部件存在、内部目标及唯一关系；保留原始部件名称/关系 ID。读取后恢复源流位置且不关闭；XML 禁止 DTD/外部解析，定期检查取消，沿用既有 ZIP/XML 资源限制。标准路径仅作为损坏/旧模板的兼容发现依据，不能替代关系图。

验证：两个 Provider 分别读取自身真实生成模板；重命名为非标准部件路径并完整改写关系/ContentTypes；覆盖共享缓存、缺失部件、错误/外部关系、重复名称/ID、取消和流所有权。全部断言清单，不能只断言发现数量。

依赖：现有完整缓存。风险：合法 OPC 路径无法解析、损坏关系被误认为普通模板。验收：实际依赖被完整定位，默认拒绝保持明确，不能以发现器存在宣告模板功能可用。

## P2 / PTPL-002：公共快照与执行预检

目标：新入口可描述保留及更新，构建失败不污染请求。

步骤：实现上述成员及批准清单，深快照并验证模式/格式/模板/名称互斥。模板读取后再验证目标、来源、物理起点及更新区域占用；请求内新建、已有保留与更新均进入统一名称预检。更新只定位精确名称/Sheet，不根据位置猜测。

验证：Abstractions/共享合同双 TFM；旧构造和 Fluent 调用；集合修改、重复渲染、失败后 Builder 和请求状态隔离。MiniExcel/SpreadCheetah 写入及枚举前明确拒绝新能力，ExcelDataReader 不声明导出支持。

风险：入口虽存在却不进主链。验收：Provider 真实消费新 SPI，未经主链验证不声明能力。

## P3 / PTPL-003：保留及更新输出

目标：两个 Provider 保存原生透视与完整来源/缓存，保留未更新透视的定义、关系及未知配置。

步骤：在引擎加载前生成模板不可变清单及磁盘快照；业务 Sheet 写入复用现有入口。保留模式保全已有 Pivot/Cache 部件、关系及未覆盖目标单元格，不依赖引擎无证据的往返承诺。输出包合并依据关系与 ContentTypes，避免重复 ID/路径；保留未涉及部件字节内容，不全包读入内存。

更新复用现有来源校验及预算，在工作簿/目标写入前验证全部更新；按实际名称更新来源范围，缓存保留全部源行和类型，重建 field.items、轴与筛选索引。共享缓存须识别所有消费者：仅更新部分消费者时拆出新缓存并分配不冲突 cacheId/关系，其它消费者继续引用原缓存；整组同源同字段更新可共享新缓存。来源扩缩不得遗留旧行或以当前筛选裁剪缓存。

轴/值/筛选等显式更新定义覆盖对应字段，未涉及布局和外观属性保留；未知扩展保持原文。若某未知扩展依赖将被改变的字段索引或外部连接而不能证明安全，预检明确拒绝该更新，不能删除扩展伪装保真。不放宽已有宏、图表和外部连接能力边界；相关透视图、切片器和时间线分别识别并固化支持/拒绝原因。

不预写静态聚合，按实际缓存及客户端刷新合同验证显示。原目标透视正文旧值与新定义的处理必须明确，不能将旧缓存显示称作更新完成；刷新配置和首开/显式刷新分别验证。资源、DOM 准入、真实外围异步 IO、取消、异常观察、默认/自定义提交器及临时文件清理保持原合同。

验证：同源/异源多透视、部分共享缓存更新、来源扩缩、保留未知 XML/局部名称/格式/行列/合并、模板输入位置与所有权；错误/取消/提交失败保持目标及输入完整。两 Provider、同步/异步真实文件职责覆盖完整结构及缓存逐值读回。

依赖：P1/P2。风险：引擎重写丢失部件、未知索引失效、旧正文冒充新值、多透视污染。验收：保留和更新分别进入公开主链，关系/内容/完整来源缓存及资源合同通过。

## P4 / PTPL-004：业务与客户端

采购/仓库模板包括基础 Sum、高级多轴五聚合、已设置数字格式/筛选的多个透视。覆盖重复请求、模板源增加/减少、缓存共享拆分、未知配置保留及更新失败重试。

用本任务包生成客户端夹具：Excel 独立隐藏实例验证首开、更新来源、显式刷新、保存重开和全部结构/数值；不得将诊断脚本退出 0 当作值通过。F-002 Count 仍未解决时保留完整失败，不缩小五聚合或改 Count 来源；环境项单独保持。资源探针仅在共享热路径实际改变时增加定向数据，不运行无关 500K/1M 矩阵。

依赖：P3。风险：包结构正确但客户端需要修复、显示陈旧或筛选失效。验收：完整值与结构一致，未验证客户端不外推。

## P5 / PTPL-005：交付

中文 XML、能力矩阵、模板迁移示例、成员级审批、双 TFM API 差异、最终生产符号到测试方法追溯、execution 和独立 Review。只批准本阶段新增成员，无删除或签名变化；候选源码/程序集/包身份重新绑定。

职责 → 两 Provider/共享合同及相关集成双 TFM → 新 NuGet 普通/第三方双 TFM 消费者 → API → solution Release 与 Category!=Large → Docs 示例 → 编码/BOM/EOL/EOF/git diff --check → 最终独立复核。使用仓库真实 dotnet/build/ApiSnapshot 入口、--no-restore -m:1 -p:UseSharedCompilation=false，不复用旧候选结果冒充当前通过。

存在 MUST_FIX/SHOULD_FIX 保持 IN_PROGRESS；F-002/003 与 F-005～007、环境目标不因本阶段局部通过宣告完成。不修改外部 Utopa、不自动通知、提交、推送、发布或版本变更。
