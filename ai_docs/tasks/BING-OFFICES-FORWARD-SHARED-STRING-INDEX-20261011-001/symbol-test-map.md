# 生产符号到测试映射

| 生产符号 | 关键行为 | 测试项目 | 测试方法 |
| --- | --- | --- | --- |
| `ExcelForwardXlsxSharedStrings.ReadDiskBacked` | 短/空/Unicode/富文本索引、随机读取与按条目定长磁盘索引 | NPOI、ClosedXML Provider 测试 | `ExcelForwardXlsxSharedStringsTest.ReadDiskBacked_ShouldKeepIndexOnDiskAndSupportRandomReads` |
| `ExcelForwardXlsxSharedStrings.ReadDiskBacked` | 预取消与 XML 异常后的双临时文件清理 | NPOI、ClosedXML Provider 测试 | `ExcelForwardXlsxSharedStringsTest.ReadDiskBacked_WhenCancelled_ShouldDeleteTemporaryFiles`; `ExcelForwardXlsxSharedStringsTest.ReadDiskBacked_WhenXmlIsInvalid_ShouldDeleteTemporaryFiles` |
| `ExcelForwardXlsxSharedStrings.ReadDiskBacked` | 空共享字符串表在预取消时仍抛出取消异常且不创建临时文件 | NPOI、ClosedXML Provider 测试 | `ExcelForwardXlsxSharedStringsTest.ReadDiskBacked_WhenPreCanceledAndEmpty_ShouldThrowBeforeCreatingFiles` |
| `ExcelForwardXlsxSharedStrings.ReadDiskBacked` | 按固定字符缓冲分块写入长共享字符串，并保留跨块 Unicode 代理对 | NPOI、ClosedXML Provider 测试 | `ExcelForwardXlsxSharedStringsTest.ReadDiskBacked_WhenStringSpansChunks_ShouldPreserveUnicodeValue` |
| `ExcelForwardXlsxReader.ReadSharedStrings` | Forward 摘要导入时使用磁盘共享字符串索引并完成关系校验 | NPOI Provider 测试 | `EntityForwardRelationTest.ImportEntityBatches_Forward_SummaryOnly_ShouldValidateRelationsWithoutRetainingEntities` |
| `ExcelForwardXlsxReader.ReadSharedStrings` | Forward 摘要导入时使用磁盘共享字符串索引并完成关系校验 | ClosedXML Provider 测试 | `EntityForwardRelationTest.ImportEntityBatches_Forward_SummaryOnly_ShouldValidateRelationsWithoutRetainingEntities` |
