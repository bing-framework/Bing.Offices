# Workbook 分批导入跨 Sheet 关系

## 目标

补齐 ExcelDataReader Workbook 分批导入的跨 Sheet 父子关系能力。现有 `ExcelWorkbookImportBuilder.HasMany`、`ExcelRelationRequest` 和完整 Workbook 导入关系语义保持不变；本轮只扩展 ExcelDataReader 的 Workbook 同步/异步分批路径，不新增公共 API。

## 行为约定

- 关系请求在预检阶段允许通过；单 Sheet 分批、完整 Import、Failure Workbook 和 Entity Layout 的既有边界保持不变。
- 分批读取期间创建 Provider 内部的 Workbook 聚合根。每个成功交付的实体在调用方回调前加入对应 Sheet 目标集合，失败行不加入。
- 所有选中 Sheet 读取完成后统一绑定关系，因此父 Sheet、子 Sheet 的物理顺序可以任意，父批次尚未交付时子批次仍可先交付。
- 关系绑定只对已成功物化的实体执行；父项缺失或父项行转换/校验失败时，子实体仍按原批次交付，关系错误在最终以子 Sheet 名称和子实体类型交付的空数据诊断批次中报告。
- 关系错误使用 `ExcelImportErrorCode.Relationship`，保留关系键为 `RawValue`；关系级错误没有可靠的单元格坐标，行列保持 0。
- 关系阶段发生在资源预算、错误预算或读取中止之前已完成的部分之后；若 Workbook 因资源上限或错误上限提前停止，则不执行关系绑定，沿用现有资源/截断结果。
- 同步和异步使用同一协调语义；异步只等待回调，不引入 `Task.Run`、`.Result` 或 `.Wait()`。
- 关系导航沿用现有 `ICollection<T>` 合同、首个父项优先、指定比较器和取消边界。

## 实施步骤

1. 新增 ExcelDataReader 内部批次关系协调器，维护聚合根、目标集合、关系索引和最终错误分组；支持跨 Sheet、乱序、父项失败和空集合。
2. 在 Workbook 同步/异步批次读取路径中接入协调器：允许关系请求、保存成功实体、完成读取后绑定并交付关系诊断批次，修正摘要错误计数。
3. 新增职责级测试：关系成功、子 Sheet 先于父 Sheet、父行失败、缺失父项、同步/异步回调顺序和取消；保留既有不支持边界测试。
4. 更新 ExcelDataReader Provider 能力文档及生产符号到测试方法追溯。

## 验收

- ExcelDataReader net6.0/net8.0 Release 构建无警告、无错误。
- 关系专项和现有 ExcelDataReader `Category!=Large` 测试通过；同步与异步结构化错误一致。
- 现有动态列、Workbook 原生校验、唯一性、资源预算和流所有权回归通过。
- 公共 API snapshot 成员差异为零；不新增生产友元。
- 执行 `git diff --check`，并检查改动 C# 文件 UTF-8 BOM、LF、末尾换行。
