using System.ComponentModel;

namespace Bing.Offices.Providers;

/// <summary>
/// 单个输入范围内的唯一值跟踪器；当前行只写入 pending journal，成功后才提交。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class UniqueTracker
{
    /// <summary>
    /// 按唯一键保存已提交的值集合，生命周期由调用方拥有。
    /// </summary>
    private readonly IDictionary<string, HashSet<string>> _committed;
    /// <summary>
    /// 按唯一键保存当前行尚未提交的值集合。
    /// </summary>
    private readonly IDictionary<string, HashSet<string>> _pending;
    /// <summary>
    /// 按唯一键和值保存已提交值首次出现的正行号。
    /// </summary>
    private readonly IDictionary<string, IDictionary<string, int>> _firstRows;
    /// <summary>
    /// 按唯一键和值保存当前行待提交值的首次正行号。
    /// </summary>
    private readonly IDictionary<string, IDictionary<string, int>> _pendingFirstRows;
    /// <summary>
    /// 回收后可复用的待提交值集合栈，减少逐行分配。
    /// </summary>
    private readonly Stack<HashSet<string>> _pendingValueBuffers;
    /// <summary>
    /// 回收后可复用的待提交首行号字典栈，减少逐行分配。
    /// </summary>
    private readonly Stack<IDictionary<string, int>> _pendingFirstRowBuffers;
    /// <summary>
    /// 允许跟踪的已提交值和待提交值总数上限；为 null 表示不限制。
    /// </summary>
    private readonly int? _maxTrackedValues;
    /// <summary>
    /// 用于唯一值集合和首行号字典的字符串比较器。
    /// </summary>
    private readonly IEqualityComparer<string> _comparer;
    /// <summary>
    /// 已提交唯一值的总数量。
    /// </summary>
    private int _trackedValueCount;
    /// <summary>
    /// 当前行待提交唯一值的数量。
    /// </summary>
    private int _pendingValueCount;

    /// <summary>
    /// 初始化一个 <see cref="UniqueTracker" /> 类型的实例。
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
        _pendingValueBuffers = new Stack<HashSet<string>>();
        _pendingFirstRowBuffers = new Stack<IDictionary<string, int>>();
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
    /// 尝试获取指定唯一值首次提交的行号。
    /// </summary>
    /// <param name="key">唯一值所属的字段键。</param>
    /// <param name="value">待查询的唯一值。</param>
    /// <param name="rowNumber">输出值首次提交的一基行号。</param>
    /// <returns>找到已提交值时为 <see langword="true" />，否则为 <see langword="false" />。</returns>
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
        ClearPending();
    }

    /// <summary>
    /// 尝试为当前行保留唯一值。
    /// </summary>
    /// <param name="key">唯一值所属的字段键。</param>
    /// <param name="value">待保留的唯一值。</param>
    /// <param name="ignoreEmpty">是否忽略空字符串和空白字符串。</param>
    /// <returns>值可被保留或按策略忽略时为 <see langword="true" />，已重复时为 <see langword="false" />。</returns>
    /// <exception cref="InvalidOperationException">跟踪值数量达到上限时抛出。</exception>
    public bool TryReserve(string key, string value, bool ignoreEmpty = true)
        => TryReserve(key, value, value == null, ignoreEmpty, 0);

    /// <summary>
    /// 尝试为当前行保留唯一值。
    /// </summary>
    /// <remarks>
    /// 在提供正数源行号时，同时记录非空值的首次行号。
    /// </remarks>
    /// <param name="key">唯一值所属的字段键。</param>
    /// <param name="value">待保留的唯一值。</param>
    /// <param name="ignoreNull">是否忽略 <see langword="null" /> 值。</param>
    /// <param name="ignoreEmpty">是否忽略空字符串和空白字符串。</param>
    /// <param name="rowNumber">值所在的一基源行号；不需要记录时传 0。</param>
    /// <returns>值可被保留或按策略忽略时为 <see langword="true" />，已重复时为 <see langword="false" />。</returns>
    /// <exception cref="InvalidOperationException">跟踪值数量达到上限时抛出。</exception>
    public bool TryReserve(string key, string value, bool ignoreNull, bool ignoreEmpty, int rowNumber)
    {
        if ((value == null && ignoreNull) || (value != null && value.Length == 0 && ignoreEmpty)
            || (value != null && string.IsNullOrWhiteSpace(value) && ignoreEmpty))
            return true;
        if (_committed.TryGetValue(key, out var committedValues) && committedValues.Contains(value))
            return false;
        var hasPending = _pending.TryGetValue(key, out var pendingValues);
        if (hasPending && pendingValues.Contains(value))
            return false;
        if (_maxTrackedValues.HasValue && _trackedValueCount + _pendingValueCount >= _maxTrackedValues.Value)
            throw new InvalidOperationException($"Unique 跟踪值超过最大数量: {_maxTrackedValues.Value}");
        if (!hasPending)
            _pending[key] = pendingValues = RentPendingValues();
        if (pendingValues.Add(value))
            _pendingValueCount++;
        if (rowNumber > 0 && value != null)
        {
            if (!_pendingFirstRows.TryGetValue(key, out var rows))
                _pendingFirstRows[key] = rows = RentPendingFirstRows();
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
            _pendingFirstRows.TryGetValue(pair.Key, out var pendingRows);
            IDictionary<string, int> firstRows = null;
            foreach (var value in pair.Value)
            {
                if (committedValues.Add(value))
                    _trackedValueCount++;
                if (value != null && pendingRows != null
                    && pendingRows.TryGetValue(value, out var rowNumber))
                {
                    if (firstRows == null && !_firstRows.TryGetValue(pair.Key, out firstRows))
                        _firstRows[pair.Key] = firstRows = new Dictionary<string, int>(_comparer);
                    if (!firstRows.ContainsKey(value))
                        firstRows[value] = rowNumber;
                }
            }
        }
        ClearPending();
    }

    /// <summary>
    /// 回滚当前行的唯一值。
    /// </summary>
    public void RollbackRow()
    {
        ClearPending();
    }

    /// <summary>
    /// 从回收池获取当前行的待提交唯一值集合。
    /// </summary>
    /// <returns>可复用的待提交唯一值集合。</returns>
    private HashSet<string> RentPendingValues()
    {
        if (_pendingValueBuffers.Count == 0)
            return new HashSet<string>(_comparer);
        return _pendingValueBuffers.Pop();
    }

    /// <summary>
    /// 从回收池获取当前行的首次行号集合。
    /// </summary>
    /// <returns>可复用的值到行号映射。</returns>
    private IDictionary<string, int> RentPendingFirstRows()
    {
        if (_pendingFirstRowBuffers.Count == 0)
            return new Dictionary<string, int>(_comparer);
        return _pendingFirstRowBuffers.Pop();
    }

    /// <summary>
    /// 清理当前行未提交的唯一值并回收到对象池。
    /// </summary>
    private void ClearPending()
    {
        foreach (var pair in _pending)
        {
            pair.Value.Clear();
            _pendingValueBuffers.Push(pair.Value);
        }
        _pending.Clear();

        foreach (var pair in _pendingFirstRows)
        {
            pair.Value.Clear();
            _pendingFirstRowBuffers.Push(pair.Value);
        }
        _pendingFirstRows.Clear();
        _pendingValueCount = 0;
    }
}
