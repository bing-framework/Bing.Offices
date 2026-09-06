# Public API 台账

## 当前分类

| 类别 | 类型数 | 结论 |
| --- | ---: | --- |
| User API | 82 | 逐项保留/收敛 |
| Provider User API | 6 | 必须保持 public |
| Provider SPI | 9 | 最小化并使用 `EditorBrowsable(Never)` |
| Execution detail | 83 | Phase 2 逐项 internal/private/移动或批准 |
| Compatibility | 0 | 不得重新引入 |

初始总计 180。Phase 2 已删除 DataTable CSV、旧 Mapping 三类型、四个 helper 与六个默认 validation 实现；最新精确分类以 `PublicApiContractTest` 为准，正式成员 hash 尚待生成候选快照。

已批准的最小跨程序集 SPI 包括 Mapping merger/registry/plan interfaces、UniqueTracker、异常 dispatcher、默认 validation 工厂与共享日期 validation rule；均必须 `EditorBrowsable(Never)`。其余可配置 DTO、Attributes、Options、方向 Builder 和扩展入口应归 User API/Provider User API，不得继续标作 Execution detail。

## 当前精确分类

| 类别 | 类型数 |
| --- | ---: |
| User API | 141 |
| Provider User API | 9 |
| Provider SPI | 15 |
| Compatibility | 0 |
| Execution detail | 0 |

合计 165 个导出类型。三个 Metadata DTO 因构成 NPOI 扩展签名归 Provider User API；其余六项为要求保留的 NPOI 扩展容器。

## 必须保持 public 的 NPOI 容器

- `CellExtensions`
- `CellStyleExtensions`
- `FontExtensions`
- `RowExtensions`
- `SheetExtensions`
- `WorkbookExtensions`
- `ExcelNpoiServiceCollectionExtensions`

当前七个容器均已导出；最终必须通过仅 PackageReference 的真实 nupkg consumer 编译和运行验证。

## Baseline 状态

正式内嵌 hash 与当前 Abstractions 不匹配，三个 Unit TFM 均失败。禁止在完成成员级 diff、Breaking Change 说明与审批证据前更新 hash。
