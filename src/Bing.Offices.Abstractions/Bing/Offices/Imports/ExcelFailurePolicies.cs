using System;
using System.IO;

namespace Bing.Offices.Imports;

/// <summary>
/// 导入失败工作簿输出模式。
/// </summary>
public enum ExcelImportFailureWorkbookMode
{
    /// <summary>不生成失败工作簿。</summary>
    None,
    /// <summary>在原工作簿副本上标记错误。</summary>
    AnnotatedOriginal,
    /// <summary>只输出包含失败行的工作簿。</summary>
    ErrorRowsOnly
}

/// <summary>
/// 失败工作簿输出边界的结构化诊断。
/// </summary>
public sealed class ExcelImportFailureDiagnostic
{
    /// <summary>创建失败工作簿诊断。</summary>
    public ExcelImportFailureDiagnostic(string code, string temporaryPath, Exception exception)
    {
        Code = code;
        TemporaryPath = temporaryPath;
        Exception = exception;
    }

    /// <summary>诊断代码。</summary>
    public string Code { get; }

    /// <summary>未包含工作簿内容的临时文件路径。</summary>
    public string TemporaryPath { get; }

    /// <summary>清理异常。</summary>
    public Exception Exception { get; }
}

/// <summary>
/// 导入失败批注与原有批注冲突时的处理策略。
/// </summary>
public enum ExcelImportCommentConflictPolicy
{
    /// <summary>保留已有批注，不追加失败信息。</summary>
    Preserve,
    /// <summary>在已有批注后追加失败信息。</summary>
    Append,
    /// <summary>用失败信息替换已有批注。</summary>
    Replace,
    /// <summary>存在已有批注时直接失败。</summary>
    Fail
}

/// <summary>
/// 导入规则来源组合策略。
/// </summary>
public enum ExcelImportValidationMode
{
    /// <summary>禁用配置和工作簿原生规则。</summary>
    Disabled,
    /// <summary>只执行配置和属性规则。</summary>
    ConfiguredRules,
    /// <summary>只执行工作簿原生规则。</summary>
    WorkbookRules,
    /// <summary>同时执行配置和工作簿规则。</summary>
    ConfiguredAndWorkbook
}

/// <summary>
/// 图片列多图片处理策略。
/// </summary>
public enum ExcelImageMultiplicityPolicy
{
    /// <summary>只绑定第一张图片。</summary>
    First,
    /// <summary>绑定全部图片。</summary>
    All,
    /// <summary>出现多张图片时报告错误。</summary>
    Fail
}

/// <summary>
/// 不支持的工作簿特性处理策略。
/// </summary>
public enum ExcelUnsupportedFeaturePolicy
{
    /// <summary>报告为导入错误。</summary>
    Report,
    /// <summary>直接拒绝导入。</summary>
    Fail
}

