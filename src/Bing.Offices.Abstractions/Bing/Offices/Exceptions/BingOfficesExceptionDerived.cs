using System;

namespace Bing.Offices.Exceptions;

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

