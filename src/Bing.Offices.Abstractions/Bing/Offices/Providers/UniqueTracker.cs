using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Bing.Offices.Providers;

/// <summary>
/// 单个输入范围内的唯一值跟踪器；当前行只写入 pending journal，成功后才提交。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class UniqueTracker
{
    private readonly IDictionary<string, HashSet<string>> _committed;
    private readonly IDictionary<string, HashSet<string>> _pending;
    private readonly IDictionary<string, IDictionary<string, int>> _firstRows;
    private readonly IDictionary<string, IDictionary<string, int>> _pendingFirstRows;
    private readonly int? _maxTrackedValues;
    private readonly IEqualityComparer<string> _comparer;
    private int _trackedValueCount;

    /// <summary>
    /// 初始化唯一值跟踪器。
    /// </summary>
    /// <param name="committed">按唯一键保存的已提交值集合。</param>
    /// <param name="maxTrackedValues">可跟踪的唯一值上限。</param>
    /// <param name="comparer">值比较器。</param>
    public UniqueTracker(IDictionary<string, HashSet<string>> committed,
        int? maxTrackedValues = null, IEqualityComparer<string> comparer = null)
    {
        _committed = committed ?? throw new ArgumentNullException(nameof(committed));
        _pending = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        _firstRows = new Dictionary<string, IDictionary<string, int>>(StringComparer.OrdinalIgnoreCase);
        _pendingFirstRows = new Dictionary<string, IDictionary<string, int>>(StringComparer.OrdinalIgnoreCase);
        _maxTrackedValues = maxTrackedValues;
        _comparer = comparer ?? StringComparer.OrdinalIgnoreCase;
        foreach (var pair in committed)
            _trackedValueCount += pair.Value?.Count ?? 0;
    }

    /// <summary>
    /// 获取已跟踪的唯一值数量。
    /// </summary>
    public int TrackedValueCount => _trackedValueCount;

    /// <summary>
    /// 获取指定唯一值首次提交的行号；不存在时返回 false。
    /// </summary>
    public bool TryGetFirstRowNumber(string key, string value, out int rowNumber)
    {
        rowNumber = 0;
        return value != null && _firstRows.TryGetValue(key, out var values)
            && values.TryGetValue(value, out rowNumber);
    }

    /// <summary>
    /// 开始新行并清除上一行未提交状态。
    /// </summary>
    public void BeginRow()
    {
        _pending.Clear();
        _pendingFirstRows.Clear();
    }

    /// <summary>
    /// 尝试为当前行保留唯一值。
    /// </summary>
    public bool TryReserve(string key, string value, bool ignoreEmpty = true)
        => TryReserve(key, value, value == null, ignoreEmpty, 0);

    /// <summary>
    /// 尝试为当前行保留唯一值，并记录首次行号。
    /// </summary>
    public bool TryReserve(string key, string value, bool ignoreNull, bool ignoreEmpty, int rowNumber)
    {
        if ((value == null && ignoreNull) || (value != null && value.Length == 0 && ignoreEmpty)
            || (value != null && string.IsNullOrWhiteSpace(value) && ignoreEmpty))
            return true;
        if (_committed.TryGetValue(key, out var committedValues) && committedValues.Contains(value))
            return false;
        if (_pending.TryGetValue(key, out var pendingValues) && pendingValues.Contains(value))
            return false;
        if (!_pending.TryGetValue(key, out pendingValues))
            _pending[key] = pendingValues = new HashSet<string>(_comparer);
        if (_maxTrackedValues.HasValue && _trackedValueCount + PendingCount() >= _maxTrackedValues.Value)
            throw new InvalidOperationException($"Unique 跟踪值超过最大数量: {_maxTrackedValues.Value}");
        pendingValues.Add(value);
        if (rowNumber > 0 && value != null)
        {
            if (!_pendingFirstRows.TryGetValue(key, out var rows))
                _pendingFirstRows[key] = rows = new Dictionary<string, int>(_comparer);
            if (!rows.ContainsKey(value))
                rows[value] = rowNumber;
        }
        return true;
    }

    /// <summary>
    /// 提交当前行的唯一值。
    /// </summary>
    public void CommitRow()
    {
        foreach (var pair in _pending)
        {
            if (!_committed.TryGetValue(pair.Key, out var committedValues))
                _committed[pair.Key] = committedValues = new HashSet<string>(_comparer);
            foreach (var value in pair.Value)
            {
                if (committedValues.Add(value))
                    _trackedValueCount++;
                if (value != null && _pendingFirstRows.TryGetValue(pair.Key, out var pendingRows)
                    && pendingRows.TryGetValue(value, out var rowNumber))
                {
                    if (!_firstRows.TryGetValue(pair.Key, out var firstRows))
                        _firstRows[pair.Key] = firstRows = new Dictionary<string, int>(_comparer);
                    if (!firstRows.ContainsKey(value))
                        firstRows[value] = rowNumber;
                }
            }
        }
        _pending.Clear();
        _pendingFirstRows.Clear();
    }

    /// <summary>
    /// 回滚当前行的唯一值。
    /// </summary>
    public void RollbackRow()
    {
        _pending.Clear();
        _pendingFirstRows.Clear();
    }

    private int PendingCount()
    {
        var count = 0;
        foreach (var values in _pending.Values)
            count += values.Count;
        return count;
    }
}
