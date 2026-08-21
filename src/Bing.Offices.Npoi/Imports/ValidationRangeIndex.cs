using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// 按行区间保存 Workbook Data Validation，查询时只访问可能覆盖目标行的区间节点。
/// </summary>
internal sealed class ValidationRangeIndex
{
    private readonly ValidationRangeNode _root;

    private ValidationRangeIndex(ValidationRangeNode root) => _root = root;

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

    internal IReadOnlyList<IDataValidation> Get(int row, int column)
    {
        return Get(row, column, out _);
    }

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
        private ValidationRangeNode(int center, IReadOnlyList<ValidationRangeEntry> overlaps,
            ValidationRangeNode left, ValidationRangeNode right)
        {
            Center = center;
            Overlaps = ValidationRangeColumnNode.Build(overlaps);
            Left = left;
            Right = right;
        }

        private int Center { get; }
        private ValidationRangeColumnNode Overlaps { get; }
        private ValidationRangeNode Left { get; }
        private ValidationRangeNode Right { get; }

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
        private ValidationRangeColumnNode(int center, IReadOnlyList<ValidationRangeEntry> overlaps,
            ValidationRangeColumnNode left, ValidationRangeColumnNode right)
        {
            Center = center;
            OverlapsByStart = overlaps.OrderBy(entry => entry.FirstColumn).ToArray();
            OverlapsByEnd = overlaps.OrderByDescending(entry => entry.LastColumn).ToArray();
            Left = left;
            Right = right;
        }

        private int Center { get; }
        private IReadOnlyList<ValidationRangeEntry> OverlapsByStart { get; }
        private IReadOnlyList<ValidationRangeEntry> OverlapsByEnd { get; }
        private ValidationRangeColumnNode Left { get; }
        private ValidationRangeColumnNode Right { get; }

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
        internal ValidationRangeEntry(int firstRow, int lastRow, int firstColumn, int lastColumn,
            IDataValidation validation)
        {
            FirstRow = firstRow;
            LastRow = lastRow;
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
            Validation = validation;
        }

        internal int FirstRow { get; }
        internal int LastRow { get; }
        internal int FirstColumn { get; }
        internal int LastColumn { get; }
        internal IDataValidation Validation { get; }

        internal bool Contains(int row, int column) => row >= FirstRow && row <= LastRow
            && column >= FirstColumn && column <= LastColumn;
    }
}
