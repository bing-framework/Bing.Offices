using System;
using System.Collections.Generic;
using Bing.Offices.Providers;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 唯一值字符额度及行级状态测试。
/// </summary>
public sealed class UniqueTrackerCharacterBudgetTest
{
    /// <summary>
    /// 测试 - 字符额度允许精确达到边界，并在超限时保留当前行已有值。
    /// </summary>
    [Fact]
    public void MaxTrackedCharacters_ShouldAllowLimitAndRejectOverflow()
    {
        var committed = new Dictionary<string, HashSet<string>>();
        var tracker = new UniqueTracker(committed, null, null, 3);

        tracker.BeginRow();
        Assert.True(tracker.TryReserve("Code", "ab", false, false, 1));
        Assert.True(tracker.TryReserve("Code", "c", false, false, 1));
        Assert.Throws<InvalidOperationException>(
            () => tracker.TryReserve("Code", "d", false, false, 1));

        tracker.CommitRow();

        Assert.Equal(2, committed["Code"].Count);
        Assert.Contains("ab", committed["Code"]);
        Assert.Contains("c", committed["Code"]);
        Assert.Throws<InvalidOperationException>(
            () => tracker.TryReserve("Code", "e", false, false, 2));
    }

    /// <summary>
    /// 测试 - 重复值不会重复消耗字符额度，超限失败不会污染待提交状态。
    /// </summary>
    [Fact]
    public void MaxTrackedCharacters_DuplicateAndFailure_ShouldPreservePendingState()
    {
        var committed = new Dictionary<string, HashSet<string>>();
        var tracker = new UniqueTracker(committed, null, StringComparer.OrdinalIgnoreCase, 3);

        tracker.BeginRow();
        Assert.True(tracker.TryReserve("Code", "ab", false, false, 1));
        Assert.False(tracker.TryReserve("Code", "AB", false, false, 1));
        Assert.Throws<InvalidOperationException>(
            () => tracker.TryReserve("Other", "xy", false, false, 1));

        tracker.CommitRow();

        Assert.Single(committed);
        Assert.Contains("ab", committed["Code"]);
        Assert.False(committed.ContainsKey("Other"));
        Assert.Throws<InvalidOperationException>(
            () => tracker.TryReserve("Other", "xy", false, false, 2));
    }

    /// <summary>
    /// 测试 - 回滚释放当前行预留的字符额度。
    /// </summary>
    [Fact]
    public void RollbackRow_ShouldReleasePendingCharacterBudget()
    {
        var committed = new Dictionary<string, HashSet<string>>();
        var tracker = new UniqueTracker(committed, null, null, 4);

        tracker.BeginRow();
        Assert.True(tracker.TryReserve("Code", "abc", false, false, 1));
        tracker.RollbackRow();

        Assert.Empty(committed);
        tracker.BeginRow();
        Assert.True(tracker.TryReserve("Code", "wxyz", false, false, 2));
        tracker.CommitRow();

        Assert.Contains("wxyz", committed["Code"]);
        Assert.Throws<InvalidOperationException>(
            () => tracker.TryReserve("Code", "q", false, false, 3));
    }

    /// <summary>
    /// 测试 - 忽略的空值不占字符额度，显式跟踪的 null 和空字符串字符长度为零。
    /// </summary>
    [Fact]
    public void TryReserve_EmptyValues_ShouldFollowIgnoreFlagsWithoutConsumingCharacters()
    {
        var committed = new Dictionary<string, HashSet<string>>();
        var tracker = new UniqueTracker(committed, null, null, 0);

        tracker.BeginRow();
        Assert.True(tracker.TryReserve("Code", null));
        Assert.True(tracker.TryReserve("Code", string.Empty));
        Assert.True(tracker.TryReserve("Code", null, false, false, 1));
        Assert.True(tracker.TryReserve("Code", string.Empty, false, false, 1));
        Assert.False(tracker.TryReserve("Code", string.Empty, false, false, 1));
        tracker.CommitRow();

        Assert.Equal(2, committed["Code"].Count);
        Assert.Throws<InvalidOperationException>(
            () => tracker.TryReserve("Code", "a", false, false, 2));
    }

    /// <summary>
    /// 测试 - 初始化时的预填值同时计入数量和字符额度。
    /// </summary>
    [Fact]
    public void Constructor_ShouldCountPrepopulatedValuesAndCharacters()
    {
        var committed = new Dictionary<string, HashSet<string>>
        {
            ["Code"] = new HashSet<string> { "ab", "c", null }
        };
        var tracker = new UniqueTracker(committed, 4, null, 3);

        Assert.Equal(3, tracker.TrackedValueCount);
        Assert.False(tracker.TryReserve("Code", "ab", false, false, 1));
        Assert.False(tracker.TryReserve("Code", null, false, false, 1));
        Assert.Throws<InvalidOperationException>(
            () => tracker.TryReserve("Other", "d", false, false, 1));
    }

    /// <summary>
    /// 测试 - 四参数构造函数仍执行数量额度，旧构造函数保持无字符额度的兼容语义。
    /// </summary>
    [Fact]
    public void Constructors_ShouldPreserveValueLimitAndLegacyCharacterBehavior()
    {
        var limited = new UniqueTracker(
            new Dictionary<string, HashSet<string>>(), 1, StringComparer.Ordinal, 100);
        Assert.True(limited.TryReserve("Code", "A", false, false, 1));
        Assert.Throws<InvalidOperationException>(
            () => limited.TryReserve("Code", "B", false, false, 2));

        var legacy = new UniqueTracker(
            new Dictionary<string, HashSet<string>>(), maxTrackedValues: 1,
            comparer: StringComparer.Ordinal);
        Assert.True(legacy.TryReserve("Code", new string('x', 128), false, false, 1));
        Assert.Throws<InvalidOperationException>(
            () => legacy.TryReserve("Code", "y", false, false, 2));
    }

    /// <summary>
    /// 测试 - 字符额度拒绝负数配置。
    /// </summary>
    [Fact]
    public void Constructor_NegativeCharacterLimit_ShouldThrow()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new UniqueTracker(new Dictionary<string, HashSet<string>>(), null, null, -1));
    }

    /// <summary>
    /// 测试 - 字符额度按照 .NET 字符串长度统计代理项对。
    /// </summary>
    [Fact]
    public void MaxTrackedCharacters_ShouldUseStringLength()
    {
        var tracker = new UniqueTracker(
            new Dictionary<string, HashSet<string>>(), null, null, 1);

        Assert.Throws<InvalidOperationException>(
            () => tracker.TryReserve("Code", "\U0001F600", false, false, 1));
    }
}
