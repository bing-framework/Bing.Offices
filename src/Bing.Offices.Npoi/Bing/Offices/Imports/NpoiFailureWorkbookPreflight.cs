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
    internal static void Validate(IWorkbook source,
        IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<string, ExcelSheetImportRequest> resolvedSheetRequests,
        ExcelImportFailureOptions options)
    {
        if (options.MaxCandidateErrorRows.HasValue && errors.Count > options.MaxCandidateErrorRows.Value)
            throw CreateLimitException($"失败工作簿候选错误行超过限制: {options.MaxCandidateErrorRows.Value}");

        if (options.Mode != ExcelImportFailureWorkbookMode.ErrorRowsOnly)
            return;

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

    private static BingOfficesResourceLimitException CreateLimitException(string message) =>
        new(message, provider: "NPOI", operation: BingOfficesOperation.Import,
            stage: BingOfficesStage.Preflight);

    /// <summary>
    /// 只统计 drawing shape 数量，避免预算检查为了读取图片数据而先分配 PictureInfo 和字节数组。
    /// </summary>
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

    private readonly struct PictureBudgetEstimate
    {
        internal PictureBudgetEstimate(int count, long bytes)
        {
            Count = count;
            Bytes = bytes;
        }

        internal int Count { get; }
        internal long Bytes { get; }
    }
}
