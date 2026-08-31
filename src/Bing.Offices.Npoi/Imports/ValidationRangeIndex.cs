using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// 按行区间保存 Workbook Data Validation，查询时只访问可能覆盖目标行的区间节点。
/// </summary>
internal sealed class ValidationRangeIndex
{
    /// <summary>按行区间组织的根节点。</summary>
    private readonly ValidationRangeNode _root;

    /// <summary>使用行区间索引根节点创建查询索引。</summary>
    /// <param name="root">已构建的行区间根节点。</param>
    private ValidationRangeIndex(ValidationRangeNode root) => _root = root;

    /// <summary>将工作簿校验区域裁剪到目标行列范围并构建索引。</summary>
    /// <param name="validations">NPOI 提供的工作簿校验规则。</param>
    /// <param name="firstRow">允许索引的最小零基行号。</param>
    /// <param name="lastRow">允许索引的最大零基行号。</param>
    /// <param name="firstColumn">允许索引的最小零基列号。</param>
    /// <param name="lastColumn">允许索引的最大零基列号。</param>
    /// <returns>可按单元格坐标查询的校验索引。</returns>
    internal static ValidationRangeIndex Create(IEnumerable<IDataValidation> validations, int firstRow, int lastRow,
        int firstColumn, int lastColumn)
    {
        var entries = new List<ValidationRangeEntry>();
        foreach (var validation in validations)
        {
            foreach (var range in validation.Regions.CellRangeAddresses)
            {
                var startRow = System.Math.Max(firstRow, range.FirstRow);
                var endRow = System.Math.Min(lastRow, range.LastRow);
                var startColumn = System.Math.Max(firstColumn, range.FirstColumn);
                var endColumn = System.Math.Min(lastColumn, range.LastColumn);
                if (startRow <= endRow && startColumn <= endColumn)
                    entries.Add(new ValidationRangeEntry(startRow, endRow, startColumn, endColumn, validation));
            }
        }
        return new ValidationRangeIndex(ValidationRangeNode.Build(entries));
    }

    /// <summary>获取覆盖指定单元格的工作簿校验规则。</summary>
    /// <param name="row">目标单元格的零基行号。</param>
    /// <param name="column">目标单元格的零基列号。</param>
    /// <returns>按索引命中的校验规则集合。</returns>
    internal IReadOnlyList<IDataValidation> Get(int row, int column)
    {
        return Get(row, column, out _);
    }

    /// <summary>获取覆盖指定单元格的校验规则并返回候选节点检查数量。</summary>
    /// <param name="row">目标单元格的零基行号。</param>
    /// <param name="column">目标单元格的零基列号。</param>
    /// <param name="candidateChecks">返回实际检查的区间候选数量。</param>
    /// <returns>按索引命中的校验规则集合。</returns>
    internal IReadOnlyList<IDataValidation> Get(int row, int column, out int candidateChecks)
    {
        candidateChecks = 0;
        if (_root == null)
            return System.Array.Empty<IDataValidation>();
        var result = new List<IDataValidation>();
        var seen = new HashSet<IDataValidation>();
        _root.Collect(row, column, result, seen, ref candidateChecks);
        return result.Count == 0 ? System.Array.Empty<IDataValidation>() : result;
    }

    private sealed class ValidationRangeNode
    {
        /// <summary>创建包含行区间重叠项及左右子树的节点。</summary>
        /// <param name="center">当前节点的中心行号。</param>
        /// <param name="overlaps">覆盖中心行的区间集合。</param>
        /// <param name="left">中心行左侧的子树。</param>
        /// <param name="right">中心行右侧的子树。</param>
        private ValidationRangeNode(int center, IReadOnlyList<ValidationRangeEntry> overlaps,
            ValidationRangeNode left, ValidationRangeNode right)
        {
            Center = center;
            Overlaps = ValidationRangeColumnNode.Build(overlaps);
            Left = left;
            Right = right;
        }

        /// <summary>获取当前节点的中心行号。</summary>
        private int Center { get; }
        /// <summary>获取按列索引的中心行重叠区间。</summary>
        private ValidationRangeColumnNode Overlaps { get; }
        /// <summary>获取中心行左侧子树。</summary>
        private ValidationRangeNode Left { get; }
        /// <summary>获取中心行右侧子树。</summary>
        private ValidationRangeNode Right { get; }

        /// <summary>递归构建行区间树。</summary>
        /// <param name="entries">待索引的校验区间。</param>
        /// <returns>构建后的节点；没有区间时为 null。</returns>
        internal static ValidationRangeNode Build(IReadOnlyList<ValidationRangeEntry> entries)
        {
            if (entries == null || entries.Count == 0)
                return null;

            var center = entries.OrderBy(entry => entry.FirstRow)
                .ElementAt(entries.Count / 2).FirstRow;
            var overlaps = new List<ValidationRangeEntry>();
            var left = new List<ValidationRangeEntry>();
            var right = new List<ValidationRangeEntry>();
            foreach (var entry in entries)
            {
                if (entry.LastRow < center)
                    left.Add(entry);
                else if (entry.FirstRow > center)
                    right.Add(entry);
                else
                    overlaps.Add(entry);
            }
            return new ValidationRangeNode(center, overlaps, Build(left), Build(right));
        }

        /// <summary>收集覆盖指定单元格的校验规则。</summary>
        /// <param name="row">目标单元格的零基行号。</param>
        /// <param name="column">目标单元格的零基列号。</param>
        /// <param name="result">接收命中规则的集合。</param>
        /// <param name="seen">防止同一规则重复加入的集合。</param>
        /// <param name="candidateChecks">累计区间候选检查次数。</param>
        internal void Collect(int row, int column, ICollection<IDataValidation> result,
            ISet<IDataValidation> seen, ref int candidateChecks)
        {
            if (row < Center)
            {
                Overlaps.Collect(column, row, result, seen, ref candidateChecks);
                Left?.Collect(row, column, result, seen, ref candidateChecks);
                return;
            }

            if (row > Center)
            {
                Overlaps.Collect(column, row, result, seen, ref candidateChecks);
                Right?.Collect(row, column, result, seen, ref candidateChecks);
                return;
            }

            Overlaps.Collect(column, row, result, seen, ref candidateChecks);
        }
    }

    private sealed class ValidationRangeColumnNode
    {
        /// <summary>创建包含列区间重叠项及左右子树的节点。</summary>
        /// <param name="center">当前节点的中心列号。</param>
        /// <param name="overlaps">覆盖中心列的区间集合。</param>
        /// <param name="left">中心列左侧的子树。</param>
        /// <param name="right">中心列右侧的子树。</param>
        private ValidationRangeColumnNode(int center, IReadOnlyList<ValidationRangeEntry> overlaps,
            ValidationRangeColumnNode left, ValidationRangeColumnNode right)
        {
            Center = center;
            OverlapsByStart = overlaps.OrderBy(entry => entry.FirstColumn).ToArray();
            OverlapsByEnd = overlaps.OrderByDescending(entry => entry.LastColumn).ToArray();
            Left = left;
            Right = right;
        }

        /// <summary>获取当前节点的中心列号。</summary>
        private int Center { get; }
        /// <summary>获取按起始列升序排列的重叠区间。</summary>
        private IReadOnlyList<ValidationRangeEntry> OverlapsByStart { get; }
        /// <summary>获取按结束列降序排列的重叠区间。</summary>
        private IReadOnlyList<ValidationRangeEntry> OverlapsByEnd { get; }
        /// <summary>获取中心列左侧子树。</summary>
        private ValidationRangeColumnNode Left { get; }
        /// <summary>获取中心列右侧子树。</summary>
        private ValidationRangeColumnNode Right { get; }

        /// <summary>递归构建列区间树。</summary>
        /// <param name="entries">待索引的校验区间。</param>
        /// <returns>构建后的节点；没有区间时为 null。</returns>
        internal static ValidationRangeColumnNode Build(IReadOnlyList<ValidationRangeEntry> entries)
        {
            if (entries == null || entries.Count == 0)
                return null;

            var center = entries.OrderBy(entry => entry.FirstColumn)
                .ElementAt(entries.Count / 2).FirstColumn;
            var overlaps = new List<ValidationRangeEntry>();
            var left = new List<ValidationRangeEntry>();
            var right = new List<ValidationRangeEntry>();
            foreach (var entry in entries)
            {
                if (entry.LastColumn < center)
                    left.Add(entry);
                else if (entry.FirstColumn > center)
                    right.Add(entry);
                else
                    overlaps.Add(entry);
            }
            return new ValidationRangeColumnNode(center, overlaps, Build(left), Build(right));
        }

        /// <summary>收集覆盖指定行列坐标的校验规则。</summary>
        /// <param name="column">目标单元格的零基列号。</param>
        /// <param name="row">目标单元格的零基行号。</param>
        /// <param name="result">接收命中规则的集合。</param>
        /// <param name="seen">防止同一规则重复加入的集合。</param>
        /// <param name="candidateChecks">累计区间候选检查次数。</param>
        internal void Collect(int column, int row, ICollection<IDataValidation> result,
            ISet<IDataValidation> seen, ref int candidateChecks)
        {
            if (column < Center)
            {
                foreach (var entry in OverlapsByStart)
                {
                    if (entry.FirstColumn > column)
                        break;
                    AddIfMatching(entry, row, column, result, seen, ref candidateChecks);
                }
                Left?.Collect(column, row, result, seen, ref candidateChecks);
                return;
            }

            if (column > Center)
            {
                foreach (var entry in OverlapsByEnd)
                {
                    if (entry.LastColumn < column)
                        break;
                    AddIfMatching(entry, row, column, result, seen, ref candidateChecks);
                }
                Right?.Collect(column, row, result, seen, ref candidateChecks);
                return;
            }

            foreach (var entry in OverlapsByStart)
                AddIfMatching(entry, row, column, result, seen, ref candidateChecks);
        }

        /// <summary>检查区间是否覆盖目标坐标，并在首次命中时加入结果。</summary>
        /// <param name="entry">待检查的校验区间。</param>
        /// <param name="row">目标零基行号。</param>
        /// <param name="column">目标零基列号。</param>
        /// <param name="result">接收命中规则的集合。</param>
        /// <param name="seen">防止同一规则重复加入的集合。</param>
        /// <param name="candidateChecks">累计区间候选检查次数。</param>
        private static void AddIfMatching(ValidationRangeEntry entry, int row, int column,
            ICollection<IDataValidation> result, ISet<IDataValidation> seen, ref int candidateChecks)
        {
            candidateChecks++;
            if (entry.Contains(row, column) && seen.Add(entry.Validation))
                result.Add(entry.Validation);
        }
    }

    private sealed class ValidationRangeEntry
    {
        /// <summary>创建一个包含闭合行列边界的校验区间条目。</summary>
        /// <param name="firstRow">最小零基行号。</param>
        /// <param name="lastRow">最大零基行号。</param>
        /// <param name="firstColumn">最小零基列号。</param>
        /// <param name="lastColumn">最大零基列号。</param>
        /// <param name="validation">区间对应的工作簿校验规则。</param>
        internal ValidationRangeEntry(int firstRow, int lastRow, int firstColumn, int lastColumn,
            IDataValidation validation)
        {
            FirstRow = firstRow;
            LastRow = lastRow;
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
            Validation = validation;
        }

        /// <summary>获取最小零基行号。</summary>
        internal int FirstRow { get; }
        /// <summary>获取最大零基行号。</summary>
        internal int LastRow { get; }
        /// <summary>获取最小零基列号。</summary>
        internal int FirstColumn { get; }
        /// <summary>获取最大零基列号。</summary>
        internal int LastColumn { get; }
        /// <summary>获取区间对应的工作簿校验规则。</summary>
        internal IDataValidation Validation { get; }

        /// <summary>判断区间是否覆盖指定行列坐标。</summary>
        /// <param name="row">目标零基行号。</param>
        /// <param name="column">目标零基列号。</param>
        /// <returns>坐标在闭合区间内时为 true。</returns>
        internal bool Contains(int row, int column) => row >= FirstRow && row <= LastRow
            && column >= FirstColumn && column <= LastColumn;
    }
}
