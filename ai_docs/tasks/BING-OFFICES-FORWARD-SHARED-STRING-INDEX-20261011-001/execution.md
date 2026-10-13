# 执行记录

**Task ID：** `BING-OFFICES-FORWARD-SHARED-STRING-INDEX-20261011-001`

## 范围

执行 `plan.md` 定义的 Forward `SummaryOnly` 共享字符串磁盘索引改造。本子项不关闭父级 F-006。

## 当前进度

- [x] 确认现有索引列表及 `ToArray()` 造成随字符串条目数增长的托管留存和复制峰值。
- [x] 完成定长磁盘索引实现与内部测试目录入口。
- [x] 增加共享职责测试，覆盖大量短/空/多语言字符串、随机读取、枚举、索引大小及异常清理。
- [x] 运行 NPOI/ClosedXML 双 TFM 直接测试、Forward `SummaryOnly` 回归和 Provider Release 构建。
- [x] 执行差异、UTF-8 BOM、LF、末尾换行和 `git diff --check` 检查。

## 验证结果

验证在当前工作区的隔离副本中执行，以避开共享 `output/debug` 中间目录的并行写入。

| 验证项 | 结果 |
| --- | --- |
| 共享索引直接测试：NPOI `net6.0`、`net8.0` | 各 3/3 通过 |
| 共享索引直接测试：ClosedXML `net6.0`、`net8.0` | 各 3/3 通过 |
| Forward `SummaryOnly` 导入回归：NPOI `net6.0`、`net8.0` | 各 4/4 通过 |
| Forward `SummaryOnly` 导入回归：ClosedXML `net6.0`、`net8.0` | 各 2/2 通过 |
| NPOI Release Provider 构建：双 TFM | 通过 |
| ClosedXML Release Provider 构建：双 TFM | 通过 |

构建出现既有 Pivot XML 文档参数警告、ClosedXML 可空注释警告及 xUnit2013 警告；无构建错误。此次验证未运行完整解决方案测试。

## 变更边界

- 未新增公共 API 或依赖。
- 未修改 CI、API snapshot、版本号。
- 未运行或依赖 Excel 客户端验收。
- 父级 F-006 仍为部分完成；本子项只消除了共享字符串条目数对应的托管偏移元数据，不覆盖 DOM、实体集合、唯一值索引或错误集合的端到端峰值预算。

## 2026-10-11 补充：空共享字符串表预取消

- `ReadDiskBacked` 在创建临时文件前检查取消令牌；空 `<sst/>` 不再因没有 `<si>` 节点而绕过取消并成功返回。
- 新增 `ReadDiskBacked_WhenPreCanceledAndEmpty_ShouldThrowBeforeCreatingFiles`，同时保留非空预取消和 XML 错误清理测试。
- 共享职责测试：NPOI、ClosedXML 的 `net6.0`、`net8.0` 各 4/4 通过；两家 Provider Release 双 TFM 构建通过。
- 未新增公共 API、依赖、CI 或 API snapshot；F-006 的端到端预算仍未完成。

## 2026-10-11 补充：共享字符串正文分块写入

- `ReadDiskBacked` 使用固定字符和字节缓冲读取 XML 文本，直接写入正文临时文件，不再为单条共享字符串创建完整 `StringBuilder`。
- 复用 UTF-8 `Encoder` 跨文本块维护代理对状态；每条字符串仍以原有定长偏移/长度记录索引，解析取消继续在块边界检查。
- 新增 `ReadDiskBacked_WhenStringSpansChunks_ShouldPreserveUnicodeValue`，验证跨 4096 字符缓冲边界的 emoji、长文本 UTF-8 字节数和释放清理。
- NPOI 与 ClosedXML 的 `net6.0`、`net8.0` 直接测试各 5/5 通过；Provider Release 编译随测试通过。存在离线 NuGet 漏洞源警告及既有 XML/xUnit 警告。
- 未新增公共 API、依赖、CI 或 API snapshot；本项减少共享字符串单值解析峰值，不代表 F-006 端到端预算完成。
