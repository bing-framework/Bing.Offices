using System;
using System.Collections.Generic;
using Bing.Offices.Providers;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// UniqueTracker 的行级提交、回滚和唯一值边界测试。
/// </summary>
public sealed class UniqueTrackerTest
{
    /// <summary>
    /// 测试 - 已提交值和当前行 pending 值均应拒绝重复值。
    /// </summary>
    [Fact]
    public void TryReserve_ShouldRejectCommittedAndPendingDuplicates()
    {
        var committed = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Code"] = new HashSet<string>(new[] { "A" }, StringComparer.OrdinalIgnoreCase)
        };
        var tracker = new UniqueTracker(committed, comparer: StringComparer.OrdinalIgnoreCase);

        tracker.BeginRow();

        Assert.False(tracker.TryReserve("Code", "a", false, false, 2));
        Assert.True(tracker.TryReserve("Code", "B", false, false, 2));
        Assert.False(tracker.TryReserve("Code", "b", false, false, 2));

        tracker.CommitRow();

        Assert.Equal(2, tracker.TrackedValueCount);
        tracker.BeginRow();
        Assert.False(tracker.TryReserve("Code", "B", false, false, 3));
    }

    /// <summary>
    /// 测试 - 回滚应丢弃 pending 值和首次行号，后续行可重新保留该值。
    /// </summary>
    [Fact]
    public void RollbackRow_ShouldDiscardPendingValuesAndFirstRows()
    {
        var tracker = new UniqueTracker(new Dictionary<string, HashSet<string>>());

        tracker.BeginRow();
        Assert.True(tracker.TryReserve("Code", "A", false, false, 4));
        tracker.RollbackRow();

        Assert.Equal(0, tracker.TrackedValueCount);
        Assert.False(tracker.TryGetFirstRowNumber("Code", "A", out _));

        tracker.BeginRow();
        Assert.True(tracker.TryReserve("Code", "A", false, false, 5));
        tracker.CommitRow();

        Assert.True(tracker.TryGetFirstRowNumber("Code", "A", out var firstRow));
        Assert.Equal(5, firstRow);
        Assert.False(tracker.TryReserve("Code", "A", false, false, 6));
    }

    /// <summary>
    /// 测试 - 首行号为第一数据行时应提交并在重复校验中保持不变。
    /// </summary>
    [Fact]
    public void CommitRow_ShouldRecordFirstDataRowNumber()
    {
        var tracker = new UniqueTracker(new Dictionary<string, HashSet<string>>());

        tracker.BeginRow();
        Assert.True(tracker.TryReserve("Code", "A", false, false, 1));
        tracker.CommitRow();

        Assert.True(tracker.TryGetFirstRowNumber("Code", "A", out var firstRow));
        Assert.Equal(1, firstRow);

        tracker.BeginRow();
        Assert.False(tracker.TryReserve("Code", "A", false, false, 2));
        Assert.True(tracker.TryGetFirstRowNumber("Code", "A", out firstRow));
        Assert.Equal(1, firstRow);
    }

    /// <summary>
    /// 测试 - 上限允许精确达到，新增值超出时抛出异常并保留已保留值。
    /// </summary>
    [Fact]
    public void MaxTrackedValues_ShouldAllowLimitAndRejectOverflow()
    {
        var tracker = new UniqueTracker(
            new Dictionary<string, HashSet<string>>(), maxTrackedValues: 2);

        tracker.BeginRow();
        Assert.True(tracker.TryReserve("Code", "A", false, false, 1));
        Assert.True(tracker.TryReserve("Code", "B", false, false, 1));
        Assert.Throws<InvalidOperationException>(
            () => tracker.TryReserve("Code", "C", false, false, 1));
        Assert.Equal(0, tracker.TrackedValueCount);

        tracker.CommitRow();

        Assert.Equal(2, tracker.TrackedValueCount);
        Assert.False(tracker.TryReserve("Code", "A", false, false, 2));
    }

    /// <summary>
    /// 测试 - 全新 key 超限失败后，提交或回滚不能把失败状态写入 tracker。
    /// </summary>
    [Fact]
    public void MaxTrackedValues_NewKeyFailure_ShouldNotPolluteCommittedOrFirstRows()
    {
        var committed = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var tracker = new UniqueTracker(committed, maxTrackedValues: 1);

        tracker.BeginRow();
        Assert.True(tracker.TryReserve("Code", "A", false, false, 10));
        tracker.CommitRow();

        tracker.BeginRow();
        Assert.Throws<InvalidOperationException>(
            () => tracker.TryReserve("Other", "B", false, false, 20));

        tracker.CommitRow();

        Assert.Equal(1, tracker.TrackedValueCount);
        Assert.Single(committed);
        Assert.True(committed.ContainsKey("Code"));
        Assert.False(committed.ContainsKey("Other"));
        Assert.True(tracker.TryGetFirstRowNumber("Code", "A", out var firstRow));
        Assert.Equal(10, firstRow);
        Assert.False(tracker.TryGetFirstRowNumber("Other", "B", out _));

        tracker.BeginRow();
        Assert.Throws<InvalidOperationException>(
            () => tracker.TryReserve("Other", "B", false, false, 21));
        tracker.RollbackRow();
        Assert.Equal(1, tracker.TrackedValueCount);
        Assert.False(committed.ContainsKey("Other"));
    }

    /// <summary>
    /// 测试 - 默认忽略 null/空值，显式关闭忽略后才参与唯一值跟踪。
    /// </summary>
    [Fact]
    public void TryReserve_ShouldFollowNullAndEmptyFlags()
    {
        var committed = new Dictionary<string, HashSet<string>>();
        var tracker = new UniqueTracker(committed);

        tracker.BeginRow();
        Assert.True(tracker.TryReserve("Code", null));
        Assert.True(tracker.TryReserve("Code", string.Empty));
        Assert.True(tracker.TryReserve("Code", " "));
        tracker.CommitRow();

        Assert.Equal(0, tracker.TrackedValueCount);
        Assert.Empty(committed);

        tracker.BeginRow();
        Assert.True(tracker.TryReserve("Code", null, false, false, 1));
        Assert.True(tracker.TryReserve("Code", string.Empty, false, false, 1));
        Assert.True(tracker.TryReserve("Code", " ", false, false, 1));
        tracker.CommitRow();

        Assert.Equal(3, tracker.TrackedValueCount);
        Assert.False(tracker.TryGetFirstRowNumber("Code", null, out _));
        Assert.True(tracker.TryGetFirstRowNumber("Code", string.Empty, out var emptyRow));
        Assert.Equal(1, emptyRow);
    }
}
