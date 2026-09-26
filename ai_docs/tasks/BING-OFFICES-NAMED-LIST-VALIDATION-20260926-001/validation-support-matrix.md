# 名称列表校验边界

对象：ClosedXML Workbook 原生 List 导入；使用既有 WorkbookRules / ConfiguredAndWorkbook 和 UnsupportedFeaturePolicy。不扩展公共导出定义。

| 形式 | 预期结果 |
| --- | --- |
| 工作簿名称，引用绝对有界单列/单行 | 按普通允许值判断 |
| 当前 Sheet 局部名称 | 局部优先于同名全局 |
| 局部名称无效但同名全局有效 | Unsupported，不回退 |
| 跨 Sheet、Sheet 名含空格/转义单引号 | 解析真实 Sheet |
| 名称指向一个绝对单元格 | 支持 |
| 未定义名称、失效 Sheet | Unsupported |
| 外部文件名称定义 | ClosedXML 名称解析器拒绝，未进入本校验管线；不承诺 Report/Fail |
| 动态公式、别名链、循环名称 | Unsupported，不求值或递归 |
| 联合区域、二维名称、整行/整列、相对名称 | Unsupported |
| 名称区域含公式单元格 | Unsupported，不触发计算；不因早期匹配绕过 |
| 空值 | 复用 IgnoreBlanks |
| 支持规则中的非法值 | WorkbookValidation，拒绝实体 |
| 不支持规则 | Report 保留实体并记录错误；Fail 拒绝实体 |
| 同步/异步 | 分类、实体和错误位置一致；源流归调用方 |

直接 A1 区域和显式列表的原有语义保留；不因新增名称功能改变其他规则。

外部文件直接列表表达式仍由既有 ClosedXmlProviderTest 覆盖 Unsupported。外部名称定义无法通过引擎的定义名解析器构造，不能将直接表达式证据等同于外部名称的 Report/Fail 证据。

验证状态以本任务 execution.md 的最终职责测试结果为准；上一任务矩阵属于该阶段历史，不作为本任务新增能力证据。
