using System;

namespace Bing.Offices.Exceptions;

/// <summary>错误码。</summary>
public enum BingOfficesErrorCode
{
    /// <summary>配置无效。</summary>
    ConfigurationInvalid,
    /// <summary>导入失败。</summary>
    ImportFailed,
    /// <summary>导出失败。</summary>
    ExportFailed,
    /// <summary>资源限制超出。</summary>
    ResourceLimitExceeded,
    /// <summary>文件提交失败。</summary>
    FileCommitFailed,
    /// <summary>不支持的功能。</summary>
    UnsupportedFeature,
    /// <summary>用户扩展执行失败。</summary>
    UserExtensionFailed
}

/// <summary>业务操作类型。</summary>
public enum BingOfficesOperation
{
    /// <summary>配置加载或解析。</summary>
    Configuration,
    /// <summary>导入。</summary>
    Import,
    /// <summary>导出。</summary>
    Export,
    /// <summary>文件提交。</summary>
    FileCommit
}

/// <summary>业务操作阶段。</summary>
public enum BingOfficesStage
{
    /// <summary>打开输入。</summary>
    Open,
    /// <summary>资源预检。</summary>
    Preflight,
    /// <summary>配置解析或映射计划。</summary>
    Plan,
    /// <summary>读取。</summary>
    Read,
    /// <summary>转换或校验。</summary>
    Validate,
    /// <summary>写入。</summary>
    Write,
    /// <summary>序列化。</summary>
    Serialize,
    /// <summary>提交。</summary>
    Commit,
    /// <summary>清理。</summary>
    Cleanup
}

/// <summary>接收 Bing.Offices 公共运行异常的观察器。</summary>
public interface IBingOfficesExceptionObserver
{
    /// <summary>观察一个已经完成分类的公共异常。</summary>
    /// <param name="exception">公共异常。</param>
    void Observe(BingOfficesException exception);
}

/// <summary>Bing.Offices 公共运行异常基类。</summary>
public class BingOfficesException : InvalidOperationException
{
    /// <summary>初始化公共运行异常。</summary>
    public BingOfficesException(BingOfficesErrorCode code, BingOfficesOperation operation,
        string provider, BingOfficesStage stage, string message, Exception innerException = null,
        string sheetName = null, int? rowIndex = null, int? columnIndex = null, string propertyName = null)
        : base(message, innerException)
    {
        Code = code;
        Operation = operation;
        Provider = provider;
        Stage = stage;
        SheetName = sheetName;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        PropertyName = propertyName;
    }

    /// <summary>获取稳定错误码。</summary>
    public BingOfficesErrorCode Code { get; }

    /// <summary>获取操作类型。</summary>
    public BingOfficesOperation Operation { get; }

    /// <summary>获取提供程序名称。</summary>
    public string Provider { get; }

    /// <summary>获取操作阶段。</summary>
    public BingOfficesStage Stage { get; }

    /// <summary>获取工作表名称。</summary>
    public string SheetName { get; }

    /// <summary>获取一基行号。</summary>
    public int? RowIndex { get; }

    /// <summary>获取一基列号。</summary>
    public int? ColumnIndex { get; }

    /// <summary>获取属性名称。</summary>
    public string PropertyName { get; }
}

/// <summary>映射、Profile 或请求配置无效异常。</summary>
public sealed class BingOfficesConfigurationException : BingOfficesException
{
    /// <summary>初始化配置异常。</summary>
    public BingOfficesConfigurationException(string message, Exception innerException = null,
        BingOfficesStage stage = BingOfficesStage.Plan)
        : base(BingOfficesErrorCode.ConfigurationInvalid, BingOfficesOperation.Configuration,
            "Core", stage, message, innerException)
    {
    }
}

/// <summary>导入公共边界不可恢复失败异常。</summary>
public sealed class BingOfficesImportException : BingOfficesException
{
    /// <summary>初始化导入异常。</summary>
    public BingOfficesImportException(string message, Exception innerException = null,
        string provider = "Core", BingOfficesStage stage = BingOfficesStage.Read,
        string sheetName = null, int? rowIndex = null, int? columnIndex = null,
        string propertyName = null, BingOfficesErrorCode code = BingOfficesErrorCode.ImportFailed)
        : base(code, BingOfficesOperation.Import, provider, stage, message, innerException,
            sheetName, rowIndex, columnIndex, propertyName)
    {
    }
}

/// <summary>导出公共边界不可恢复失败异常。</summary>
public sealed class BingOfficesExportException : BingOfficesException
{
    /// <summary>初始化导出异常。</summary>
    public BingOfficesExportException(string message, Exception innerException = null,
        string provider = "Core", BingOfficesStage stage = BingOfficesStage.Write,
        string sheetName = null, int? rowIndex = null, int? columnIndex = null,
        string propertyName = null, BingOfficesErrorCode code = BingOfficesErrorCode.ExportFailed)
        : base(code, BingOfficesOperation.Export, provider, stage, message, innerException,
            sheetName, rowIndex, columnIndex, propertyName)
    {
    }
}

/// <summary>输入或输出资源预算超出异常。</summary>
public sealed class BingOfficesResourceLimitException : BingOfficesException
{
    /// <summary>初始化资源限制异常。</summary>
    public BingOfficesResourceLimitException(string message, Exception innerException = null,
        string provider = "Core", BingOfficesOperation operation = BingOfficesOperation.Import,
        BingOfficesStage stage = BingOfficesStage.Preflight)
        : base(BingOfficesErrorCode.ResourceLimitExceeded, operation, provider, stage,
            message, innerException)
    {
    }
}

/// <summary>原子文件提交异常。</summary>
public sealed class BingOfficesFileCommitException : BingOfficesException
{
    /// <summary>初始化文件提交异常。</summary>
    public BingOfficesFileCommitException(string message, Exception innerException = null,
        string provider = "Core", BingOfficesStage stage = BingOfficesStage.Commit)
        : base(BingOfficesErrorCode.FileCommitFailed, BingOfficesOperation.FileCommit,
            provider, stage, message, innerException)
    {
    }
}

/// <summary>当前提供程序不支持请求功能异常。</summary>
public sealed class BingOfficesUnsupportedFeatureException : BingOfficesException
{
    /// <summary>初始化不支持功能异常。</summary>
    public BingOfficesUnsupportedFeatureException(string message, Exception innerException = null,
        string provider = "Core", BingOfficesOperation operation = BingOfficesOperation.Import,
        BingOfficesStage stage = BingOfficesStage.Read)
        : base(BingOfficesErrorCode.UnsupportedFeature, operation, provider, stage,
            message, innerException)
    {
    }
}
