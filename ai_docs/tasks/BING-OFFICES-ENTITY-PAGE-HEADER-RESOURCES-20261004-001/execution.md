# 实施执行记录

Task ID：`BING-OFFICES-ENTITY-PAGE-HEADER-RESOURCES-20261004-001`

## 完成情况

已完成 DESIGN-004。NPOI 和 ClosedXML 的 Entity 重复页眉导入在列表末行回溯、共享 marker 扫描、首页页眉源区域检查和复制页表头核对时传递并定期检查取消令牌。保留原有行扫描与结构化错误坐标，没有增加公共成员或改动旧签名。

真实稀疏工作簿验证：XLS 与 XLSX 使用格式允许的大 `End`，ClosedXML XLSX 在第 20,000 行放置非列表数据；导入均只读取列表自身的明细。NPOI 的输入字节限制在打开工作簿阶段报资源错误；ClosedXML 的明细行数限制在第二条明细处报错。共享 marker 扫描与复制页表头核对均有运行中取消的确定性测试。没有使用易受机器负载影响的时间阈值，暂未发现需要改用“已用单元格扫描”的功能性或可执行性能回归。

## 验证

- 重复页眉职责测试：NPOI net6.0/net8.0 各 88 项；ClosedXML 各 54 项，均通过。
- `Category!=Large` Release 解决方案常规测试通过；其中 NPOI Provider 双 TFM 各 947 项，Provider 合同各 298 项，Docs 19 项。最后新增两项 ClosedXML 取消测试后，该 Provider 双 TFM 全套各 339 项通过。
- Release 解决方案构建 0 错误；串行打包成功。本地 NuGet 包消费者 net6.0、net8.0 均输出 `package-consumer-ok`。
- 双 TFM 公共 API 成员差异为零；只更新源码清单及 NPOI、ClosedXML 包身份，正式 API snapshot 比较通过。
- `git diff --check` 通过；本轮 13 个文本文件通过严格 UTF-8、BOM、LF 和末尾换行检查。`git diff --stat` 已检查，工作区包含此前任务的大量未提交改动，均予以保留。
- 保留既有 `Microsoft.Bcl.Memory` net6.0 和 SDK 生命周期提示；未修改版本号，未提交、推送或发布。

## 后续

DESIGN-005：提取共享重复页眉 Provider 合同，保留各引擎真实文件结构断言。DESIGN-006：透视表、公式计算和真实打印渲染单独设计。
