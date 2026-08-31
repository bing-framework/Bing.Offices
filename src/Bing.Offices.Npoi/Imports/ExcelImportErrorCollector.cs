using Bing.Offices.Imports;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// Workbook 级有界错误收集器。子收集器只负责 Sheet 视图，限制始终由根收集器执行。
/// </summary>
internal sealed class ExcelImportErrorCollector
{
    /// <summary>父级收集器；子收集器将错误同步到该根收集器以共享上限。</summary>
    private readonly ExcelImportErrorCollector _parent;
    /// <summary>根收集器允许保留的最大错误数量。</summary>
    private readonly int? _maxErrors;
    /// <summary>当前收集器视图中的错误集合。</summary>
    private readonly List<ExcelImportError> _errors = new();

    /// <summary>使用指定的根错误上限初始化收集器。</summary>
    /// <param name="maxErrors">允许保留的最大错误数；为 null 时不限制。</param>
    internal ExcelImportErrorCollector(int? maxErrors)
    {
        _maxErrors = maxErrors;
    }

    /// <summary>创建共享根错误上限的工作表级错误视图。</summary>
    /// <param name="parent">负责全局限制的父级收集器。</param>
    private ExcelImportErrorCollector(ExcelImportErrorCollector parent)
    {
        _parent = parent;
    }

    /// <summary>创建用于单个工作表的子收集器。</summary>
    /// <returns>与当前根收集器共享限制的子收集器。</returns>
    internal ExcelImportErrorCollector CreateChild() => new(this);

    /// <summary>获取当前收集器视图可见的错误。</summary>
    internal IReadOnlyList<ExcelImportError> Errors => _parent == null ? _errors : _errors;

    /// <summary>获取根收集器当前保留的错误数量。</summary>
    internal int Count => _parent == null ? _errors.Count : _parent.Count;

    /// <summary>获取错误是否因数量上限而被截断。</summary>
    internal bool IsTruncated => _parent == null ? _isTruncated : _parent.IsTruncated;

    /// <summary>获取根收集器允许保留的最大错误数量。</summary>
    internal int? MaxErrors => _parent == null ? _maxErrors : _parent.MaxErrors;

    /// <summary>指示根收集器是否已发生错误截断。</summary>
    private bool _isTruncated;

    /// <summary>获取根收集器是否已达到错误数量上限。</summary>
    internal bool IsLimitReached => _parent == null ? _maxErrors.HasValue && _errors.Count >= _maxErrors.Value : _parent.IsLimitReached;

    /// <summary>将根收集器标记为结果已被截断。</summary>
    internal void MarkTruncated()
    {
        if (_parent != null)
        {
            _parent.MarkTruncated();
            return;
        }
        _isTruncated = true;
    }

    /// <summary>向根收集器及当前工作表视图添加一条错误。</summary>
    /// <param name="error">要记录的导入错误。</param>
    /// <returns>成功记录错误时为 true；错误为空或达到限制时为 false。</returns>
    internal bool Add(ExcelImportError error)
    {
        if (error == null)
            return false;
        if (_parent != null)
        {
            if (!_parent.Add(error))
                return false;
            _errors.Add(error);
            return true;
        }
        if (_maxErrors.HasValue && _errors.Count >= _maxErrors.Value)
        {
            _isTruncated = true;
            return false;
        }
        _errors.Add(error);
        if (_maxErrors.HasValue && _errors.Count >= _maxErrors.Value)
            _isTruncated = true;
        return true;
    }

}
