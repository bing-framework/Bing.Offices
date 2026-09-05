# Resource Report

## Excel 独立进程探针

Artifact：`ai_docs/tasks/BING-OFFICES-RC-HARDENING-20260904-001/artifacts/excel-resource-probe-rerun.jsonl`

共 12 个 child-process 场景：

- 正常/既有：`zip`、`dom`、`dom-limit`、`shared-strings`、`styles`、`drawings`、`ole`。
- 当前新增预检拒绝：`zip-total-limit`、`zip-ratio-limit`、`shared-strings-limit`、`styles-limit`、`worksheet-limit`。

5 个新增拒绝场景均满足：`status=resource-limit`、`rejectStage=Preflight`、`importedRows=0`、进程退出码 `0`。正常 `zip`、`dom`、shared strings、styles、drawings、OLE 样本均成功；`dom-limit` 继续证明行数限制为结构化资源限制，不与 ZIP preflight 混淆。

代表性观测：

- `zip`：4 source rows / 3 imported rows，peak working set `51,437,568` bytes。
- `dom`：250 source rows / 249 imported rows，peak working set `54,841,344` bytes。
- mapping/unique 100k workload 最大观察到 LOH `56,197,384` bytes，peak working set `137,875,456` bytes。

## Mapping/Unique 独立探针

Artifact：`ai_docs/tasks/BING-OFFICES-RC-HARDENING-20260904-001/artifacts/mapping-resource-probe.jsonl`

- 16/16 child scenarios：`passed`。
- LOH ceiling：`536,870,912` bytes。
- Peak working set ceiling：`1,073,741,824` bytes。
- 最大观测：LOH `56,197,384` bytes，peak working set `137,875,456` bytes。
- workload：mapping plan and unique tracker；不能替代 Excel DOM 或 Failure Workbook 资源证据。

## 安全边界

- `NpoiXlsxZipPreflight.Validate` 在 `WorkbookFactory.Create` 前执行。
- 检查 entry count、单 entry/总解压大小、压缩比、`sharedStrings.xml`、`styles.xml`、单 worksheet/worksheet 总量、重复 entry、异常路径和 XML DTD/entity。
- XML reader 使用 `DtdProcessing.Prohibit`、`XmlResolver=null`，并按 entry 长度限制文档字符数；当前没有独立最大 XML 深度预算。
- `MaxInputBytes`、ZIP 解压预算、NPOI DOM、实体对象图、Failure Workbook 序列化和进程工作集是不同边界，不能互相替代。
- XLS/OLE 没有 ZIP 等价的 DOM 前内部预检；当前只保留输入大小、独立进程和部署资源限制声明。

## 未完成门禁

- 未完成 Failure Workbook `AnnotatedOriginal`/`ErrorRowsOnly` 双 DOM 的完整峰值矩阵。
- 未完成 100k/1M Excel import/export 全矩阵和取消延迟矩阵。
- 未形成经维护者批准的 LOH/working-set/输入/输出预算。
- 当前证据证明已实现的拒绝路径和观测样本，不证明任意 Workbook 的硬内存上限。

## 结论

ZIP preflight reject evidence：`PASS`；mapping/unique probe：`16/16 PASS`；完整资源发布门禁：`PARTIAL / UNAPPROVED`。
