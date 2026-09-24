# Provider Capability Matrix

| Provider | List | Workbook | Entity | Template | Merge | Async | XLSX | XLS |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| NPOI | yes | yes | yes | yes | yes | yes | yes | yes |
| MiniExcel | yes | yes | no | no | no | yes | yes | no |
| ClosedXML | yes | yes | yes | yes | template-preserving/basic | yes | yes | no |

现有 `ExcelProviderCapabilities` bit 数值未改变，也没有为尚未具备公共 requirement 的 Rich XLSX 能力追加公共 bit。样式、公式和模板结构的细粒度 preflight 保持在 Provider 内部。

ClosedXML Entity 已覆盖多个同 Sheet/跨 Sheet List Region 与 `HasMany` Relations；Entity List Region 动态列仍为 `ACCEPTED_LIMITATION`。
