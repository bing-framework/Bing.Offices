# 异常与 Observer

不可恢复的 Office 业务失败统一使用 `BingOfficesException` 家族。调用方应优先依赖 `Code`、`Operation`、`Provider` 和 `Stage` 做稳定判断；`Message` 面向诊断和本地化，不是程序合同。导入中的单元格或记录错误仍进入 `ExcelImportError` / `CsvImportError`，不提升为操作级异常。

参数错误保持 `ArgumentNullException`、`ArgumentException` 或 `ArgumentOutOfRangeException`；取消保持 `OperationCanceledException`；`OutOfMemoryException`、`StackOverflowException` 等致命异常不会被 Try API 或公共边界吞掉。用户 Converter、Validator 和关系委托失败会保留原始 `InnerException`，并在适用时携带 Sheet、行、列和属性位置。

通过依赖注入注册 `IBingOfficesExceptionObserver` 后，Excel、CSV 和 DI 配置加载公共边界会通知已完成分类的异常。Observer 收到的对象与调用方捕获的对象是同一实例，每次操作最多通知一次；Observer 自身失败只附加到主异常诊断，不覆盖主异常。静态 `ExcelMappingConfigurationLoader` 是纯解析入口，不参与 Observer；需要统一观察时应解析 `IExcelMappingConfigurationLoader`。

文件导出使用同目录临时文件。只有临时文件创建、Flush、Replace 或 Move 的文件系统失败归类为 `BingOfficesFileCommitException`；写委托中的配置、转换或序列化失败保持原分类。直接写调用方 Stream 不提供回滚保证，失败后 Stream 可能已包含部分内容。

应用日志建议至少记录异常类型、`Code`、`Operation`、`Provider`、`Stage` 和位置字段。不要只按异常消息文本分支。
