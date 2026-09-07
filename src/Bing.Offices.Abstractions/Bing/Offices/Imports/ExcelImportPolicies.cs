using System;
using System.IO;

namespace Bing.Offices.Imports;

/// <summary>导入失败工作簿和诊断输出的策略。</summary>
public sealed class ExcelImportFailureOptions
{
    /// <summary>失败工作簿模式。</summary>
    public ExcelImportFailureWorkbookMode Mode { get; init; }

    /// <summary>失败工作簿目标流，由调用方拥有。</summary>
    public Stream Destination { get; init; }

    /// <summary>失败工作簿序列化输出允许的最大字节数。</summary>
    public long? MaxSerializedBytes { get; init; }

    /// <summary>ErrorRowsOnly 模式允许复制的最大错误数据行数。</summary>
    public int? MaxCandidateErrorRows { get; init; }

    /// <summary>ErrorRowsOnly 模式允许复制的最大单元格估算数量。</summary>
    public long? MaxCopiedCells { get; init; }

    /// <summary>ErrorRowsOnly 模式允许复制的最大图片数量。</summary>
    public int? MaxCopiedPictures { get; init; }

    /// <summary>ErrorRowsOnly 模式允许复制的图片数据最大字节数。</summary>
    public long? MaxCopiedPictureBytes { get; init; }

    /// <summary>ErrorRowsOnly 模式允许目标工作簿估算对象的最大数量；估算包括 Sheet、行、单元格和图片对象。</summary>
    public long? MaxEstimatedTargetObjects { get; init; }

    /// <summary>失败工作簿请求级临时目录；为空时使用系统临时目录。</summary>
    public string TemporaryDirectory { get; init; }

    /// <summary>失败工作簿边界诊断接收器。</summary>
    public Action<ExcelImportFailureDiagnostic> DiagnosticSink { get; init; }

    /// <summary>AnnotatedOriginal 模式下的批注冲突策略，默认追加失败信息。</summary>
    public ExcelImportCommentConflictPolicy CommentConflictPolicy { get; init; } =
        ExcelImportCommentConflictPolicy.Append;

    /// <summary>验证失败输出配置。</summary>
    public void Validate()
    {
        if (MaxSerializedBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxSerializedBytes));
        if (MaxCandidateErrorRows <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxCandidateErrorRows));
        if (MaxCopiedCells <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxCopiedCells));
        if (MaxCopiedPictures <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxCopiedPictures));
        if (MaxCopiedPictureBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxCopiedPictureBytes));
        if (MaxEstimatedTargetObjects <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxEstimatedTargetObjects));
        if (Mode != ExcelImportFailureWorkbookMode.None && Destination == null)
            throw new ArgumentException("启用失败工作簿输出时必须提供目标流。", nameof(Destination));
        if (Destination != null && !Destination.CanWrite)
            throw new ArgumentException("失败工作簿目标流不可写入。", nameof(Destination));
        if (TemporaryDirectory != null && string.IsNullOrWhiteSpace(TemporaryDirectory))
            throw new ArgumentException("临时目录不能为空白字符串。", nameof(TemporaryDirectory));
        if (!Enum.IsDefined(typeof(ExcelImportCommentConflictPolicy), CommentConflictPolicy))
            throw new ArgumentOutOfRangeException(nameof(CommentConflictPolicy));
    }
}
