using Bing.Offices.Imports;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// Workbook 级有界错误收集器。子收集器只负责 Sheet 视图，限制始终由根收集器执行。
/// </summary>
internal sealed class ExcelImportErrorCollector
{
    private readonly ExcelImportErrorCollector _parent;
    private readonly int? _maxErrors;
    private readonly List<ExcelImportError> _errors = new();

    internal ExcelImportErrorCollector(int? maxErrors)
    {
        _maxErrors = maxErrors;
    }

    private ExcelImportErrorCollector(ExcelImportErrorCollector parent)
    {
        _parent = parent;
    }

    internal ExcelImportErrorCollector CreateChild() => new(this);

    internal IReadOnlyList<ExcelImportError> Errors => _parent == null ? _errors : _errors;

    internal int Count => _parent == null ? _errors.Count : _parent.Count;

    internal bool IsTruncated => _parent == null ? _isTruncated : _parent.IsTruncated;

    internal int? MaxErrors => _parent == null ? _maxErrors : _parent.MaxErrors;

    private bool _isTruncated;

    internal bool IsLimitReached => _parent == null ? _maxErrors.HasValue && _errors.Count >= _maxErrors.Value : _parent.IsLimitReached;

    internal void MarkTruncated()
    {
        if (_parent != null)
        {
            _parent.MarkTruncated();
            return;
        }
        _isTruncated = true;
    }

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
