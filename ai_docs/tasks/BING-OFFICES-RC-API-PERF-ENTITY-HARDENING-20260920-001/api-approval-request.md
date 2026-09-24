# API 成员级审批请求

## 任务与候选

- Task ID：`BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001`
- 候选 API diff：`artifacts/api/compare/current-20260921-candidate/api-diff.json`
- 当前候选 manifest：`artifacts/candidates/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/manifest-current-20260921.md`
- 当前 snapshot baseline：`build/api-snapshot-baseline.json`
- 当前 baseline 尚未更新；本记录由用户在 2026-09-21 明确批准 `FIX-001` 后形成，批准范围覆盖下表全部 API 变更。baseline 更新仍必须使用当前 immutable candidate 并通过双 TFM 门禁。

## 变更规模

| 程序集 | 新增成员 | 删除成员 |
|---|---:|---:|
| `Bing.Offices.Abstractions` | 111 | 6 |
| `Bing.Offices.Core` | 11 | 0 |
| `Bing.Offices.Npoi` | 18 | 2（类型接口形状变化） |
| `Bing.Offices.MiniExcel` | 8 | 2（类型接口形状变化） |

双 TFM 的当前候选 API canonical identity 相同；完整成员级机器差异以上述当前候选 `api-diff.json` 为准，禁止用汇总数字或历史 compare 产物替代逐成员审查。

2026-09-21 刷新说明：当前候选在 FIX-005 修复后新增的 7 个 Abstractions 成员也属于本次批准范围，具体为 `ExcelEntityCellBuilder<TValue>` 类型、无参构造函数、两个 `Mapping(...)` 重载、`ExcelEntityLayoutBuilder<TEntity>.Cell(...configure)` 重载，以及 `ExcelEntityCellBinding<TValue>.MappingConfiguration` 和 `MappingDocument` 属性。它们与既有 Entity layout/result/template 合同一并批准；当前 source-scope hash 为 `A1D558E350FF8819A60D7502684CDB4F84BE0B965F9CD56AA3B304C316355E89`。

## 删除与重命名

以下成员属于明确 Breaking 变更，需逐项批准：

1. 删除 `Bing.Offices.Imports.ValidateMode` 枚举及 `Continue`、`StopOnFirstFailure` 字段。
2. 将 `ExcelSheetImportBuilder<TItem>.Validate(ValidateMode)` 重命名为 `Validate(ExcelValidationFailureMode)`。
3. 将 `ExcelSheetImportRequest.ValidateMode` 重命名为 `ValidationFailureMode`。
4. `NpoiExcelImporter`、`NpoiExcelExporter`、`MiniExcelExcelImporter`、`MiniExcelExcelExporter` 的公开类型接口集合增加 Entity/Capability SPI；类型本身不删除，旧接口实现仍保留。

迁移示例：

```csharp
builder.Validate(ExcelValidationFailureMode.Continue);
request.ValidationFailureMode = ExcelValidationFailureMode.StopOnFirstFailure;
```

## 新增公共契约分组

以下分组中的每一个构造函数、属性、方法、枚举字段和类型都需要成员级决定；完整签名见机器 diff：

1. `Bing.Offices.Entities.ExcelEntity*`：固定 Cell、Merge、List Region、多 Sheet、Template options、ImportResult、布局构建器和值对象。
2. `Bing.Offices.Entities.IExcelEntityImporter` 与 `IExcelEntityExporter`：Entity/Template 的同步、异步、Stream 和 file 入口。
3. `Bing.Offices.Extensions.ExcelEntityExtensions`：对既有 `IExcelImporter`/`IExcelExporter` 的薄扩展入口。
4. `Bing.Offices.Providers.ExcelProviderCapabilities` 与 `IExcelProviderCapabilities`：只读能力声明和 preflight 查询，不提供运行时 Provider 路由。
5. Provider 实现的 `ProviderName`、`Capabilities`、`Supports` 以及 Entity/Template 入口；MiniExcel 对未声明能力保持结构化 fail-fast。

## 审批决定

| 范围 | 决定 | 批准人 | 时间 | 备注 |
|---|---|---|---|---|
| `ValidateMode` rename | `APPROVED` | 用户（FIX-001） | 2026-09-21 | 按当前迁移示例执行，不增加兼容 wrapper |
| Entity layout/result/template | `APPROVED` | 用户（FIX-001） | 2026-09-21 | 接受固定 Cell、Merge、List Region 和 Template 合同 |
| Entity importer/exporter SPI | `APPROVED` | 用户（FIX-001） | 2026-09-21 | 接受 Stream、取消、错误和 LeaveOpen 合同 |
| Provider capability SPI | `APPROVED` | 用户（FIX-001） | 2026-09-21 | 仅用于 preflight，不引入 router |
| NPOI/MiniExcel 公开实现接口形状 | `APPROVED` | 用户（FIX-001） | 2026-09-21 | 接受第三方 PackageReference 消费方式 |

## 批准后的执行顺序

1. 将本表所有范围改为明确的 `APPROVED` 或 `REJECTED`，拒绝项回退对应 API，不得静默保留。
2. 对每个批准的删除、重命名和新增成员补充 old→new 迁移记录。
3. 重新 capture 双 TFM API snapshot，核对 candidate identity 和版本文件哈希。
4. 仅在审批范围完全覆盖机器 diff 后更新 `build/api-snapshot-baseline.json`，并重跑 `PublicApiContractTest`。

本次用户批准已解除 API 成员审批阻断；完成 baseline 更新、双 TFM compare 和 Public API 测试后，API gate 才可标记为 `PASS`。其他性能、资源和外部环境门禁不受本批准影响。
