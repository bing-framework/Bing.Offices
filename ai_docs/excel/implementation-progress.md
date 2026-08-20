# Excel 高级能力实施进度

## 工作区保护

- 开始日期：2026-08-18
- API 策略：按用户确认，未发布 API 直接破坏性收敛，不保留兼容包装器、Obsolete 类型或平行主调用链。
- 开始实施前工作区已有大量未提交变更，包含现有 Excel 重构、删除和新增文件；本次不回退这些改动。
- 未执行 Git Commit、Push、Tag 或 PR。

## 阶段状态

| 阶段 | 状态 | 说明 |
|---|---|---|
| 0. 基线与 P0 | 已完成 | P0 回归、Setter、Culture、Duplicate、结果统计已落地 |
| 1. Workbook/Sheet 主模型 | 已完成 | Workbook Export/Import Request 支持异构 Sheet 和显式关系 |
| 2. 动态列/样式/布局 | 已完成 | typed 动态列、Key/Alias/DataType、Order/Placement、样式缓存已接入 |
| 3. 多 Sheet 导出 | 已完成 | 多 Sheet、导航集合、隐藏状态和单次枚举主链已接入 |
| 4. 多 Sheet 导入/关系 | 已完成 | 按 Sheet Header/Data 行解析，父子键绑定和错误结果已接入 |
| 5. 模板 | 已完成 | 模板来源、Sheet 缺失、命名区域锚点和流所有权已接入 |
| 6. 图表 | 已完成 | provider-neutral Chart/Series/Range 和 XLSX 柱状/折线/饼图已接入 |
| 7. 文档/基准/最终验证 | 已完成 | 文档、Benchmark、API approval、多目标构建测试和打包已完成 |

## 已确认现状

- 解决方案包含 Abstractions、Core、Npoi、Unit Test、Integration Test、Benchmark 项目。
- Abstractions/Core 目标为 `netstandard2.0`；Npoi 目标包含 `net8.0;net7.0;net6.0;netcoreapp3.1`。
- NPOI 锁定版本为 `2.7.4`。
- `IExcelExporter`/`IExcelImporter` 公开契约已收敛为 Workbook Request。
- 生产程序集已删除 `ExcelExportOptions<T>`、`ExcelImportOptions<T>` 和 `LegacyExcelFileImportOptions<T>`。
- 测试项目保留独立迁移适配器，仅用于运行历史回归，不进入生产程序集。
- 18 个历史 XLSX 资源存在，但尚未形成高级模板/图表/真实 Office 资源测试矩阵。
- Public API 当前主要由测试中的公开顶层类型清单校验，未覆盖完整成员签名。

## 基线命令与结果

1. `dotnet restore Bing.Offices.sln --locked-mode`
   - 结果：失败。
   - 原因：当前 `.csproj` 与已有 `packages.lock.json` 不一致，出现 `NU1004`。
2. `dotnet restore Bing.Offices.sln --force-evaluate`
   - 结果：通过，并重新生成当前项目依赖锁定结果。
3. `dotnet build Bing.Offices.sln -c Release --no-restore`
   - 结果：通过；9 个目标框架相关警告，主要来自 netcoreapp3.1/net5.0 使用的 System.Security 8.x 包支持声明。
4. `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net8.0 -c Release --no-restore`
   - 结果：97 通过，0 失败。
5. `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net6.0 -c Release --no-restore`
   - 结果：99 通过，0 失败。
6. `dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net8.0 -c Release --no-restore`
   - 结果：10 通过，0 失败。
7. `dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net6.0 -c Release --no-restore`
   - 结果：10 通过，0 失败。

8. `dotnet build Bing.Offices.sln -c Release --no-restore`
   - 结果：通过；包含 NPOI/旧目标框架支持警告，无编译错误。
9. `dotnet pack Bing.Offices.sln -c Release --no-build --no-restore`
   - 结果：通过；生成可打包项目产物。

## 待处理 P0

- Attribute 自定义 Enum/Nullable Enum 数值转换与目标属性赋值。
- `ColumnIndex` 的精确位置语义与动态列布局。
- decimal/int 等数值 Formatter 保持 Numeric Cell 并应用 DataFormat。
- Culture 贯穿 Range/Date 校验和错误信息。
- 明确 ValidateMode 的停止粒度以及 Continue 行为。
- 丰富 Import Result/Error 的成功状态、统计、Sheet/Row/Column/Header/RawValue/ColumnKey。
- Duplicate 使用增量提交/回滚，避免逐行深拷贝。
- Setter、Converter、Validator 计划级预绑定。
- 导出输出避免 `MemoryStream.ToArray()` 整文件复制。

## 下一步

1. 在安装 Excel/LibreOffice 的环境中补充真实模板和图表互操作验证。
2. 继续维护消费者侧迁移示例，避免重新引入单 Sheet Options API。

## 待验证项

- Excel/LibreOffice 制作的真实模板和图表互操作性尚未验证。
- 当前环境是否安装 Excel 或 LibreOffice 尚未确认。
- `.xls` 高级样式、模板动态列、图表能力需以 NPOI 2.7.4 实测结论为准。
