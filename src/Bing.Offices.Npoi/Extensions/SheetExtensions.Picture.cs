using System.Diagnostics;
using Bing.Offices.Metadata;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// NPOI工作表(<see cref="NPOI.SS.UserModel.ISheet"/>) 扩展
/// </summary>
internal static partial class SheetExtensions
{
    /// <summary>
    /// 将图片数据按指定锚点和样式添加到工作表绘图区。
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    /// <param name="picInfo">图片信息</param>
    public static void AddPicture(this NPOI.SS.UserModel.ISheet sheet, PictureInfo picInfo)
    {
        if (picInfo is null)
            throw new ArgumentNullException(nameof(picInfo));
        if (picInfo.PictureData is null || picInfo.PictureData.Length == 0)
            throw new ArgumentException("图片数据不能为空", nameof(picInfo));
        if (picInfo.PictureStyle is null)
            throw new ArgumentException("图片样式不能为空", nameof(picInfo));
        var pictureIdx = sheet.Workbook.AddPicture(picInfo.PictureData,
            PictureTypeResolver.Resolve(picInfo.PictureData));
        var anchor = sheet.Workbook.GetCreationHelper().CreateClientAnchor();
        anchor.Row1 = picInfo.MinRow;
        anchor.Row2 = picInfo.MaxRow;
        anchor.Col1 = picInfo.MinCol;
        anchor.Col2 = picInfo.MaxCol;
        anchor.Dx1 = picInfo.PictureStyle.AnchorDx1;
        anchor.Dx2 = picInfo.PictureStyle.AnchorDx2;
        anchor.Dy1 = picInfo.PictureStyle.AnchorDy1;
        anchor.Dy2 = picInfo.PictureStyle.AnchorDy2;
        var drawing = sheet.DrawingPatriarch ?? sheet.CreateDrawingPatriarch();
        drawing.CreatePicture(anchor, pictureIdx);
    }

    /// <summary>
    /// 获取工作表中所有图片的锚点、数据和样式信息。
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    /// <returns>工作表中的图片信息；没有图片时返回空列表。</returns>
    public static List<PictureInfo> GetAllPictureInfos(this NPOI.SS.UserModel.ISheet sheet) => sheet.GetAllPictureInfos(null, null, null, null);

    /// <summary>
    /// 获取工作表指定区域内或与其相交的图片信息。
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    /// <param name="minRow">最小行索引</param>
    /// <param name="maxRow">最大行索引</param>
    /// <param name="minCol">最小列索引</param>
    /// <param name="maxCol">最大列索引</param>
    /// <param name="onlyInternal">是否仅返回完全位于指定区域内的图片。</param>
    /// <returns>符合区域条件的图片信息；没有匹配项时返回空列表。</returns>
    public static List<PictureInfo> GetAllPictureInfos(this NPOI.SS.UserModel.ISheet sheet, int? minRow,
        int? maxRow, int? minCol, int? maxCol, bool onlyInternal = true)
    {
        switch (sheet)
        {
            case HSSFSheet hssfSheet:
                return GetAllPictureInfos(hssfSheet, minRow, maxRow, minCol, maxCol, onlyInternal);
            case XSSFSheet xssfSheet:
                return GetAllPictureInfos(xssfSheet, minRow, maxRow, minCol, maxCol, onlyInternal);
        }
        throw new NotImplementedException($"尚未实现该[{sheet.GetType()}]类型的[{nameof(GetAllPictureInfos)}]扩展方法");
    }

    /// <summary>
    /// 从 HSSF 工作表读取符合区域条件的图片信息。
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="minRow">最小行索引</param>
    /// <param name="maxRow">最大行索引</param>
    /// <param name="minCol">最小列索引</param>
    /// <param name="maxCol">最大列索引</param>
    /// <param name="onlyInternal">是否仅返回完全位于指定区域内的图片。</param>
    /// <returns>符合条件的图片信息列表。</returns>
    private static List<PictureInfo> GetAllPictureInfos(HSSFSheet sheet, int? minRow,
        int? maxRow, int? minCol, int? maxCol, bool onlyInternal)
    {
        var result = new List<PictureInfo>();
        if (sheet.DrawingPatriarch is HSSFShapeContainer shapeContainer)
        {
            foreach (var shape in shapeContainer.Children)
            {
                if (shape is HSSFPicture picture && picture.ClientAnchor is HSSFClientAnchor anchor)
                {
                    if (!IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2,
                            anchor.Col1,
                            anchor.Col2, onlyInternal))
                        continue;
                    var picStyle = new PictureStyle()
                    {
                        AnchorDx1 = anchor.Dx1,
                        AnchorDx2 = anchor.Dx2,
                        AnchorDy1 = anchor.Dy1,
                        AnchorDy2 = anchor.Dy2,
                        IsNoFill = picture.IsNoFill,
                        //LineStyle = picture.LineStyle,
                        LineStyleColor = picture.LineStyleColor,
                        LineWidth = picture.LineWidth,
                        FillColor = picture.FillColor,
                    };
                    result.Add(new PictureInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2,
                        picture.PictureData.Data, picStyle));
                }
            }
        }
        return result;
    }

    /// <summary>
    /// 从 XSSF 工作表读取符合区域条件的图片信息。
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="minRow">最小行索引</param>
    /// <param name="maxRow">最大行索引</param>
    /// <param name="minCol">最小列索引</param>
    /// <param name="maxCol">最大列索引</param>
    /// <param name="onlyInternal">是否仅返回完全位于指定区域内的图片。</param>
    /// <returns>符合条件的图片信息列表。</returns>
    private static List<PictureInfo> GetAllPictureInfos(XSSFSheet sheet, int? minRow,
        int? maxRow, int? minCol, int? maxCol, bool onlyInternal)
    {
        var result = new List<PictureInfo>();
        foreach (var documentPart in sheet.GetRelations())
        {
            if (documentPart is XSSFDrawing drawing)
            {
                foreach (var shape in drawing.GetShapes())
                {
                    if (shape is not XSSFPicture picture || picture.ClientAnchor == null)
                        continue;
                    var anchor = picture.ClientAnchor;
                    if (!IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2,
                            anchor.Col1,
                            anchor.Col2, onlyInternal))
                        continue;
                    var picStyle = new PictureStyle()
                    {
                        AnchorDx1 = anchor.Dx1,
                        AnchorDx2 = anchor.Dx2,
                        AnchorDy1 = anchor.Dy1,
                        AnchorDy2 = anchor.Dy2,
                    };
                    result.Add(new PictureInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2,
                        picture.PictureData.Data, picStyle));
                }
            }
        }
        return result;
    }

    /// <summary>
    /// 移除工作表中的所有图片。
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    public static void RemovePictures(this NPOI.SS.UserModel.ISheet sheet) => sheet.RemovePictures(null, null, null, null);

    /// <summary>
    /// 移除工作表指定区域内或与其相交的图片。
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    /// <param name="minRow">最小行索引</param>
    /// <param name="maxRow">最大行索引</param>
    /// <param name="minCol">最小列索引</param>
    /// <param name="maxCol">最大列索引</param>
    /// <param name="onlyInternal">是否仅移除完全位于指定区域内的图片。</param>
    public static void RemovePictures(this NPOI.SS.UserModel.ISheet sheet, int? minRow, int? maxRow, int? minCol,
        int? maxCol, bool onlyInternal = true)
    {
        switch (sheet)
        {
            case HSSFSheet hssfSheet:
                RemovePictures(hssfSheet, minRow, maxRow, minCol, maxCol, onlyInternal);
                return;
            case XSSFSheet xssfSheet:
                RemovePictures(xssfSheet, minRow, maxRow, minCol, maxCol, onlyInternal);
                return;
        }
        throw new NotImplementedException($"尚未实现该[{sheet.GetType()}]类型的[{nameof(RemovePictures)}]扩展方法");
    }

    /// <summary>
    /// 从 HSSF 工作表移除符合区域条件的图片。
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="minRow">最小行索引</param>
    /// <param name="maxRow">最大行索引</param>
    /// <param name="minCol">最小列索引</param>
    /// <param name="maxCol">最大列索引</param>
    /// <param name="onlyInternal">是否仅移除完全位于指定区域内的图片。</param>
    private static void RemovePictures(HSSFSheet sheet, int? minRow, int? maxRow, int? minCol,
        int? maxCol, bool onlyInternal)
    {
        if (!(sheet.DrawingPatriarch is HSSFShapeContainer shapeContainer))
            return;
        var pictures = shapeContainer.Children
            .OfType<HSSFPicture>()
            .Where(picture => picture.ClientAnchor is HSSFClientAnchor anchor &&
                              IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2,
                                  anchor.Col1, anchor.Col2, onlyInternal))
            .ToList();
        foreach (var picture in pictures)
        {
            shapeContainer.RemoveShape(picture);
        }
    }

    /// <summary>
    /// 从 XSSF 工作表移除符合区域条件的图片。
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="minRow">最小行索引</param>
    /// <param name="maxRow">最大行索引</param>
    /// <param name="minCol">最小列索引</param>
    /// <param name="maxCol">最大列索引</param>
    /// <param name="onlyInternal">是否仅移除完全位于指定区域内的图片。</param>
    private static void RemovePictures(XSSFSheet sheet, int? minRow, int? maxRow, int? minCol,
        int? maxCol, bool onlyInternal)
    {
        foreach (var drawing in sheet.GetRelations().OfType<XSSFDrawing>())
        {
            drawing.GetCTDrawing().CellAnchors.RemoveAll(anchor =>
                IsPictureInRange(anchor, minRow, maxRow, minCol, maxCol, onlyInternal));
        }
    }

    /// <summary>
    /// 移动工作表中的所有图片锚点。
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    /// <param name="moveRowCount">移动行数</param>
    /// <param name="moveColCount">移动列数</param>
    public static void
        MovePictures(this NPOI.SS.UserModel.ISheet sheet, int moveRowCount = 0, int moveColCount = 0) =>
        sheet.MovePictures(null, null, null, null, true, moveRowCount, moveColCount);

    /// <summary>
    /// 按行列偏移移动指定区域内的图片锚点。
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    /// <param name="minRow">最小行索引</param>
    /// <param name="maxRow">最大行索引</param>
    /// <param name="minCol">最小列索引</param>
    /// <param name="maxCol">最大列索引</param>
    /// <param name="onlyInternal">是否仅移动完全位于指定区域内的图片。</param>
    /// <param name="moveRowCount">移动行数</param>
    /// <param name="moveColCount">移动列数</param>
    public static void MovePictures(this NPOI.SS.UserModel.ISheet sheet, int? minRow, int? maxRow, int? minCol,
        int? maxCol, bool onlyInternal = true, int moveRowCount = 0, int moveColCount = 0)
    {
        switch (sheet)
        {
            case HSSFSheet hssfSheet:
                MovePictures(hssfSheet, minRow, maxRow, minCol, maxCol, onlyInternal, moveRowCount, moveColCount);
                return;
            case XSSFSheet xssfSheet:
                MovePictures(xssfSheet, minRow, maxRow, minCol, maxCol, onlyInternal, moveRowCount, moveColCount);
                return;
        }
    }

    /// <summary>
    /// 移动 HSSF 图片锚点。
    /// </summary>
    /// <param name="sheet">工作表。</param>
    /// <param name="minRow">最小行索引。</param>
    /// <param name="maxRow">最大行索引。</param>
    /// <param name="minCol">最小列索引。</param>
    /// <param name="maxCol">最大列索引。</param>
    /// <param name="onlyInternal">是否仅移动完全位于区域内的图片。</param>
    /// <param name="moveRowCount">行偏移量。</param>
    /// <param name="moveColCount">列偏移量。</param>
    private static void MovePictures(HSSFSheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol,
        bool onlyInternal, int moveRowCount, int moveColCount)
    {
        if (sheet.DrawingPatriarch is not HSSFShapeContainer shapeContainer)
            return;
        foreach (var picture in shapeContainer.Children.OfType<HSSFPicture>())
        {
            if (picture.ClientAnchor is IClientAnchor anchor)
                MovePictureAnchor(anchor, minRow, maxRow, minCol, maxCol, onlyInternal, moveRowCount, moveColCount);
        }
    }

    /// <summary>
    /// 移动 XSSF 图片锚点。
    /// </summary>
    /// <param name="sheet">工作表。</param>
    /// <param name="minRow">最小行索引。</param>
    /// <param name="maxRow">最大行索引。</param>
    /// <param name="minCol">最小列索引。</param>
    /// <param name="maxCol">最大列索引。</param>
    /// <param name="onlyInternal">是否仅移动完全位于区域内的图片。</param>
    /// <param name="moveRowCount">行偏移量。</param>
    /// <param name="moveColCount">列偏移量。</param>
    private static void MovePictures(XSSFSheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol,
        bool onlyInternal, int moveRowCount, int moveColCount)
    {
        foreach (var drawing in sheet.GetRelations().OfType<XSSFDrawing>())
        {
            foreach (var picture in drawing.GetShapes().OfType<XSSFPicture>())
            {
                if (picture.ClientAnchor is IClientAnchor anchor)
                    MovePictureAnchor(anchor, minRow, maxRow, minCol, maxCol, onlyInternal, moveRowCount,
                        moveColCount);
            }
        }
    }

    /// <summary>
    /// 原地移动图片锚点，保留图片关系、类型及样式。
    /// </summary>
    private static void MovePictureAnchor(IClientAnchor anchor, int? minRow, int? maxRow, int? minCol, int? maxCol,
        bool onlyInternal, int moveRowCount, int moveColCount)
    {
        if (!IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2, anchor.Col1,
                anchor.Col2, onlyInternal))
            return;
        anchor.Row1 = Math.Max(0, anchor.Row1 + moveRowCount);
        anchor.Row2 = Math.Max(0, anchor.Row2 + moveRowCount);
        anchor.Col1 = Math.Max(0, anchor.Col1 + moveColCount);
        anchor.Col2 = Math.Max(0, anchor.Col2 + moveColCount);
    }

    /// <summary>
    /// 判断 DrawingML 锚点是否为区域内的图片。
    /// </summary>
    /// <param name="anchor">DrawingML 锚点。</param>
    /// <param name="minRow">最小行索引。</param>
    /// <param name="maxRow">最大行索引。</param>
    /// <param name="minCol">最小列索引。</param>
    /// <param name="maxCol">最大列索引。</param>
    /// <param name="onlyInternal">是否仅匹配完全位于区域内的锚点。</param>
    /// <returns>锚点为符合区域条件的图片时返回 true。</returns>
    private static bool IsPictureInRange(IEG_Anchor anchor, int? minRow, int? maxRow, int? minCol, int? maxCol,
        bool onlyInternal)
    {
        if (anchor.picture == null || !TryGetAnchorRange(anchor, out var firstRow, out var lastRow, out var firstCol,
                out var lastCol))
            return false;
        return IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, firstRow, lastRow, firstCol, lastCol,
            onlyInternal);
    }

    /// <summary>
    /// 读取 DrawingML 锚点坐标范围。
    /// </summary>
    /// <param name="anchor">DrawingML 锚点。</param>
    /// <param name="firstRow">起始行索引。</param>
    /// <param name="lastRow">结束行索引。</param>
    /// <param name="firstCol">起始列索引。</param>
    /// <param name="lastCol">结束列索引。</param>
    /// <returns>锚点含有可读取的单元格坐标时返回 true。</returns>
    private static bool TryGetAnchorRange(IEG_Anchor anchor, out int firstRow, out int lastRow, out int firstCol,
        out int lastCol)
    {
        switch (anchor)
        {
            case CT_TwoCellAnchor twoCellAnchor:
                firstRow = twoCellAnchor.from.row;
                lastRow = twoCellAnchor.to.row;
                firstCol = twoCellAnchor.from.col;
                lastCol = twoCellAnchor.to.col;
                return true;
            case CT_OneCellAnchor oneCellAnchor:
                firstRow = lastRow = oneCellAnchor.from.row;
                firstCol = lastCol = oneCellAnchor.from.col;
                return true;
            default:
                firstRow = lastRow = firstCol = lastCol = 0;
                return false;
        }
    }

    /// <summary>
    /// 将已有 NPOI 图片数据添加到工作表并自动调整图片大小。
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="row">行索引</param>
    /// <param name="col">列索引</param>
    /// <param name="pictureData">图片数据</param>
    /// <returns>图片成功添加时为 true；NPOI 拒绝图片数据或创建绘图区失败时为 false。</returns>
    public static bool TryAddPicture(this ISheet sheet, int row, int col, IPictureData pictureData) => TryAddPicture(sheet, row, col, pictureData.Data, pictureData.PictureType);

    /// <summary>
    /// 将图片字节添加到工作表并自动调整图片大小；失败时返回 false。
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="row">行索引</param>
    /// <param name="col">列索引</param>
    /// <param name="pictureBytes">图片数据</param>
    /// <param name="pictureType">图片类型</param>
    /// <returns>图片成功添加时为 true；图片数据无效或 NPOI 创建失败时为 false。</returns>
    public static bool TryAddPicture(this ISheet sheet, int row, int col, byte[] pictureBytes,
        PictureType pictureType = PictureType.PNG)
    {
        if (sheet is null)
            throw new ArgumentNullException(nameof(sheet));

        try
        {
            var pictureIndex = sheet.Workbook.AddPicture(pictureBytes, pictureType);

            var clientAnchor = sheet.Workbook.GetCreationHelper().CreateClientAnchor();
            clientAnchor.Row1 = row;
            clientAnchor.Col1 = col;

            var picture = (sheet.DrawingPatriarch ?? sheet.CreateDrawingPatriarch())
                .CreatePicture(clientAnchor, pictureIndex);
            picture.Resize();
            return true;
        }
        catch (Exception e)
        {
            Debug.WriteLine(e);
        }
        return false;
    }
}
