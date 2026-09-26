using Bing.Offices.Exceptions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Imports;

/// <summary>
/// Failure Workbook 的资源预检职责；在创建目标工作簿前拒绝超出预算的请求。
/// </summary>
internal static class NpoiFailureWorkbookPreflight
{
    /// <summary>
    /// 在创建失败工作簿前校验候选错误、图片和目标对象资源预算。
    /// </summary>
    /// <param name="source">包含原始工作表的源工作簿。</param>
    /// <param name="errors">待写入失败工作簿的导入错误集合。</param>
    /// <param name="resolvedSheetRequests">按实际工作表名称索引的导入请求。</param>
    /// <param name="options">失败工作簿输出及资源限制选项。</param>
    internal static void Validate(IWorkbook source,
        IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<string, ExcelSheetImportRequest> resolvedSheetRequests,
        ExcelImportFailureOptions options)
    {
        if (options.Mode != ExcelImportFailureWorkbookMode.ErrorRowsOnly)
            return;

        var candidateRows = errors.Where(error => error.RowIndex > 1)
            .Select(error => (error.SheetName, error.RowIndex))
            .Distinct(new ErrorRowComparer())
            .Count();
        if (options.MaxCandidateErrorRows.HasValue && candidateRows > options.MaxCandidateErrorRows.Value)
            throw CreateLimitException($"失败工作簿候选错误行超过限制: {options.MaxCandidateErrorRows.Value}");

        var pictureEstimate = options.MaxCopiedPictures.HasValue
            || options.MaxCopiedPictureBytes.HasValue
            || options.MaxEstimatedTargetObjects.HasValue
            ? CountPictures(source)
            : default;
        if (options.MaxCopiedPictures.HasValue && pictureEstimate.Count > options.MaxCopiedPictures.Value)
            throw CreateLimitException($"失败工作簿图片数量超过限制: {options.MaxCopiedPictures.Value}");

        if (options.MaxCopiedPictureBytes.HasValue
            && pictureEstimate.Bytes > options.MaxCopiedPictureBytes.Value)
            throw CreateLimitException($"失败工作簿图片字节数超过限制: {options.MaxCopiedPictureBytes.Value}");

        if (!options.MaxCopiedCells.HasValue && !options.MaxEstimatedTargetObjects.HasValue)
            return;

        var groups = errors.Where(error => error.RowIndex > 1)
            .GroupBy(error => error.SheetName, StringComparer.OrdinalIgnoreCase);
        long estimatedCells = 0;
        long estimatedObjects = pictureEstimate.Count;
        foreach (var group in groups)
        {
            var sourceSheet = source.GetSheet(group.Key);
            if (sourceSheet == null)
                continue;
            var headerRowIndex = resolvedSheetRequests.TryGetValue(sourceSheet.SheetName, out var request)
                ? request.HeaderRowIndex
                : sourceSheet.FirstRowNum;
            var sourceRows = new HashSet<int>(group.Select(error => error.RowIndex - 1))
            {
                headerRowIndex
            };
            var groupCells = sourceRows.Sum(row => Math.Max(0,
                (int)(sourceSheet.GetRow(row)?.LastCellNum ?? 0)));
            estimatedCells += groupCells;
            estimatedObjects += 1L + sourceRows.Count + groupCells;
            if (options.MaxCopiedCells.HasValue && estimatedCells > options.MaxCopiedCells.Value)
                throw CreateLimitException($"失败工作簿复制单元格估算数量超过限制: {options.MaxCopiedCells.Value}");
            if (options.MaxEstimatedTargetObjects.HasValue
                && estimatedObjects > options.MaxEstimatedTargetObjects.Value)
                throw CreateLimitException(
                    $"失败工作簿目标对象估算数量超过限制: {options.MaxEstimatedTargetObjects.Value}");
        }
    }

    /// <summary>
    /// 创建失败工作簿预检阶段的资源限制异常。
    /// </summary>
    /// <param name="message">资源限制原因。</param>
    /// <returns>带有 NPOI 导入阶段信息的资源限制异常。</returns>
    private static BingOfficesResourceLimitException CreateLimitException(string message) =>
        new(message, provider: "NPOI", operation: BingOfficesOperation.Import,
            stage: BingOfficesStage.Preflight);

    /// <summary>
    /// 统计 drawing 中的图片数量和数据字节数，供失败工作簿资源预算使用。
    /// </summary>
    /// <param name="workbook">待统计图片的源工作簿。</param>
    /// <returns>预计处理的图片数量和字节数。</returns>
    private static PictureBudgetEstimate CountPictures(IWorkbook workbook)
    {
        var count = 0;
        long bytes = 0;
        for (var index = 0; index < workbook.NumberOfSheets; index++)
        {
            var sheet = workbook.GetSheetAt(index);
            switch (sheet)
            {
                case HSSFSheet hssfSheet:
                    if (hssfSheet.DrawingPatriarch is HSSFShapeContainer hssfContainer)
                    {
                        foreach (var picture in hssfContainer.Children.OfType<HSSFPicture>())
                        {
                            count++;
                            bytes += picture.PictureData?.Data?.LongLength ?? 0;
                        }
                    }
                    break;
                case XSSFSheet xssfSheet:
                    foreach (var picture in xssfSheet.GetRelations().OfType<XSSFDrawing>()
                                 .SelectMany(drawing => drawing.GetShapes()).OfType<XSSFPicture>())
                    {
                        count++;
                        bytes += picture.PictureData?.Data?.LongLength ?? 0;
                    }
                    break;
                default:
                    throw new NotSupportedException(
                        $"不支持该[{sheet.GetType()}]类型的图片预算预检。");
            }
        }

        return new PictureBudgetEstimate(count, bytes);
    }

    /// <summary>
    /// 保存失败工作簿图片预检得到的数量和字节数。
    /// </summary>
    private readonly struct PictureBudgetEstimate
    {
        /// <summary>
        /// 初始化一个 <see cref="PictureBudgetEstimate" /> 类型的实例。
        /// </summary>
        /// <param name="count">预计处理的图片数量。</param>
        /// <param name="bytes">预计处理的图片数据大小（字节）。</param>
        internal PictureBudgetEstimate(int count, long bytes)
        {
            Count = count;
            Bytes = bytes;
        }

        /// <summary>
        /// 获取预计处理的图片数量。
        /// </summary>
        internal int Count { get; }
        /// <summary>
        /// 获取预计处理的图片数据大小（字节）。
        /// </summary>
        internal long Bytes { get; }
    }

    /// <summary>
    /// 按 Sheet 名称（忽略大小写）和物理行号标识失败工作簿候选行。
    /// </summary>
    private sealed class ErrorRowComparer : IEqualityComparer<(string SheetName, int RowIndex)>
    {
        /// <inheritdoc />
        public bool Equals((string SheetName, int RowIndex) left, (string SheetName, int RowIndex) right) =>
            left.RowIndex == right.RowIndex
            && string.Equals(left.SheetName, right.SheetName, StringComparison.OrdinalIgnoreCase);

        /// <inheritdoc />
        public int GetHashCode((string SheetName, int RowIndex) value) =>
            HashCode.Combine(StringComparer.OrdinalIgnoreCase.GetHashCode(value.SheetName ?? string.Empty), value.RowIndex);
    }
}
