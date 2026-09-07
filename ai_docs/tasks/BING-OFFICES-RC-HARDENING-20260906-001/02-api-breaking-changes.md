# API 变更与快照报告

## 状态

`BLOCKED`：已从计划指定基线 commit 生成 before snapshot 和正式 baseline 文件，但维护者尚未填写 `approvedBy`/`approvedAt`，因此门禁仍不能 PASS。

## 生成器

- Schema：`bing.offices.public-api.v2`
- Generator：`2.0.0`
- 工具：`build/ApiSnapshot`
- 候选目录：`artifacts/api-snapshot/candidate`
- 资源预算新增后复验目录：`artifacts/api-snapshot/final-rerun-5`
- 比较命令：`dotnet build/ApiSnapshot/bin/Release/net8.0/ApiSnapshot.dll --root output/release --output .../artifacts/api-snapshot/compare --commit 958e5b4886c2fe8df80ece3218d9fab1c57a0ec6`
- 比较结果：资源预算新增后复验仍退出码 `1`，原因包含 baseline approval metadata 缺失；新增 Abstractions 预算属性已进入 candidate diff，未自动批准。

## Before 快照

| TFM | Assembly | Members | SHA-256 |
| --- | --- | ---: | --- |
| net8.0 | Bing.Offices.Abstractions | 775 | `B95FB6C3D194235A6B83C50730FD7F307E639E6BFD2A7D012159646EAAFCEB29` |
| net8.0 | Bing.Offices.Core | 152 | `28DC6D40E29E05C3F3009290494377DDAD84E50F7D4BF9A116D37B8FFEC43CA0` |
| net8.0 | Bing.Offices.Npoi | 66 | `9ABCB32CA4DA51E68B1C5004AA992EFEED776F1C82A4914143CC9DA9FFC07352` |

机器 diff：Abstractions `9 added / 5 removed`，Core `3 added / 2 removed`，Npoi `0 changed`。删除项是 diagnostics API；新增项是文件导出、提交 SPI、Failure Workbook 预算属性及新增图片字节/目标对象预算。

## 候选快照

| TFM | Assembly | Members | SHA-256 |
| --- | --- | ---: | --- |
| net8.0 | Bing.Offices.Abstractions | 779 | `3975E26F418893CA8D753BE8C959530F96FB2E815E8C7FD7B069EAB322128860` |
| net8.0 | Bing.Offices.Core | 152 | `3689ECB9C4B8FED50838C42D45CC7DDF749C27463555053019255FBD2CE69D49` |
| net8.0 | Bing.Offices.Npoi | 66 | `9ABCB32CA4DA51E68B1C5004AA992EFEED776F1C82A4914143CC9DA9FFC07352` |

Canonicalizer 已覆盖类型 kind/visibility/abstract/sealed/base/interfaces、泛型约束、成员修饰符、返回值、参数 ref/out/名称/默认值、属性 accessor、事件 accessor 和 custom attributes；测试 `PublicApiSnapshot_CanonicalLines_ShouldIncludeGovernedMetadata` 直接断言关键完整成员行。

## 已实现的公共面变化

- `IExcelExporter.ExportToFile(...)` 与 `ICsvExporter.ExportToFile<T>(...)` 成为正式 exporter 能力。
- 新增隐藏 Provider SPI：`IFileExportCommitter` 与 `DefaultFileExportCommitter`，均标记 `EditorBrowsable(Never)`。
- 删除 `ExcelMappingDiagnostic` 及 JSON/XML `out IReadOnlyList<ExcelMappingDiagnostic>` 重载和私有 diagnostics 参数。
- 新增 `MaxCopiedPictureBytes`、`MaxEstimatedTargetObjects` Failure Workbook 预算属性；维护者审批前保持 API BLOCKED。
- NPOI、Unit、Integration 的发布/测试目标收敛到 net8.0；Abstractions/Core 保持 netstandard2.0。

## 审批要求

维护者必须审阅 `api-candidate.json` 与 `api-diff.json`，填写正式 baseline 的 schema、基线 commit、审批人和审批时间后，才能解除 Unit 中被标记为 BLOCKED 的 API 门禁。
