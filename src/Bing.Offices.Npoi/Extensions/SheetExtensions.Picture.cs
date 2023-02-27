using System;
using System.Collections.Generic;
using System.Diagnostics;
using Bing.Offices.Metadata;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// NPOI工作表(<see cref="NPOI.SS.UserModel.ISheet"/>) 扩展
/// </summary>
public static partial class SheetExtensions
{
    /// <summary>
    /// 添加图片
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    /// <param name="picInfo">图片信息</param>
    public static void AddPicture(this NPOI.SS.UserModel.ISheet sheet, PictureInfo picInfo)
    {
        var pictureIdx = sheet.Workbook.AddPicture(picInfo.PictureData, PictureType.PNG);
        var anchor = sheet.Workbook.GetCreationHelper().CreateClientAnchor();
        anchor.Col1 = picInfo.MinCol;
        anchor.Col2 = picInfo.MaxCol;
        anchor.Row1 = picInfo.MinRow;
        anchor.Row2 = picInfo.MaxRow;
        anchor.Dx1 = picInfo.PictureStyle.AnchorDx1;
        anchor.Dx2 = picInfo.PictureStyle.AnchorDx2;
        anchor.Dy1 = picInfo.PictureStyle.AnchorDy1;
        anchor.Dy2 = picInfo.PictureStyle.AnchorDy2;
        anchor.AnchorType = AnchorType.MoveDontResize;
        var drawing = sheet.CreateDrawingPatriarch();
        var pic = drawing.CreatePicture(anchor, pictureIdx);
        if (sheet is HSSFSheet)
        {
            var shape = pic as HSSFShape;
            shape.FillColor = picInfo.PictureStyle.FillColor;
            shape.IsNoFill = picInfo.PictureStyle.IsNoFill;
            //shape.LineStyle = picInfo.PictureStyle.LineStyle;
            shape.LineStyleColor = picInfo.PictureStyle.LineStyleColor;
            shape.LineWidth = (int)picInfo.PictureStyle.LineWidth;
        }
        else if (sheet is XSSFSheet)
        {
            var shape = pic as XSSFShape;
            shape.FillColor = picInfo.PictureStyle.FillColor;
            shape.IsNoFill = picInfo.PictureStyle.IsNoFill;
            //shape.LineStyle = picInfo.PictureStyle.LineStyle;
            //shape.LineStyleColor = picInfo.PictureStyle.LineStyleColor;
            shape.LineWidth = picInfo.PictureStyle.LineWidth;
        }
    }

    /// <summary>
    /// 获取工作表中包含图片的信息列表
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    public static List<PictureInfo> GetAllPictureInfos(this NPOI.SS.UserModel.ISheet sheet) => sheet.GetAllPictureInfos(null, null, null, null);

    /// <summary>
    /// 获取工作表中指定区域包含图片的信息列表
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    /// <param name="minRow">最小行索引</param>
    /// <param name="maxRow">最大行索引</param>
    /// <param name="minCol">最小列索引</param>
    /// <param name="maxCol">最大列索引</param>
    /// <param name="onlyInternal">是否仅在内部</param>
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
    /// 获取所有图片信息列表
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="minRow">最小行索引</param>
    /// <param name="maxRow">最大行索引</param>
    /// <param name="minCol">最小列索引</param>
    /// <param name="maxCol">最大列索引</param>
    /// <param name="onlyInternal">是否仅在内部</param>
    private static List<PictureInfo> GetAllPictureInfos(HSSFSheet sheet, int? minRow,
        int? maxRow, int? minCol, int? maxCol, bool onlyInternal)
    {
        var result = new List<PictureInfo>();
        if (sheet.DrawingPatriarch is HSSFShapeContainer shapeContainer)
        {
            foreach (var shape in shapeContainer.Children)
            {
                if (shape is HSSFPicture picture && picture.Anchor is HSSFClientAnchor anchor)
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
    /// 获取所有图片信息列表
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="minRow">最小行索引</param>
    /// <param name="maxRow">最大行索引</param>
    /// <param name="minCol">最小列索引</param>
    /// <param name="maxCol">最大列索引</param>
    /// <param name="onlyInternal">是否仅在内部</param>
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
                    var picture = (XSSFPicture)shape;
                    var anchor = picture.GetPreferredSize();
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
    /// 移除工作表中所有的图片
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    public static void RemovePictures(this NPOI.SS.UserModel.ISheet sheet) => sheet.RemovePictures(null, null, null, null);

    /// <summary>
    /// 移除工作表中指定区域的图片
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    /// <param name="minRow">最小行索引</param>
    /// <param name="maxRow">最大行索引</param>
    /// <param name="minCol">最小列索引</param>
    /// <param name="maxCol">最大列索引</param>
    /// <param name="onlyInternal">是否仅在内部</param>
    public static void RemovePictures(this NPOI.SS.UserModel.ISheet sheet, int? minRow, int? maxRow, int? minCol,
        int? maxCol, bool onlyInternal = true)
    {
        switch (sheet)
        {
            case HSSFSheet hssfSheet:
                RemovePictures(hssfSheet, minRow, maxRow, minCol, maxCol, onlyInternal);
                break;
            case XSSFSheet xssfSheet:
                RemovePictures(xssfSheet, minRow, maxRow, minCol, maxCol, onlyInternal);
                break;
        }
        throw new NotImplementedException($"尚未实现该[{sheet.GetType()}]类型的[{nameof(RemovePictures)}]扩展方法");
    }

    /// <summary>
    /// 移除图片
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="minRow">最小行索引</param>
    /// <param name="maxRow">最大行索引</param>
    /// <param name="minCol">最小列索引</param>
    /// <param name="maxCol">最大列索引</param>
    /// <param name="onlyInternal">是否仅在内部</param>
    private static void RemovePictures(HSSFSheet sheet, int? minRow, int? maxRow, int? minCol,
        int? maxCol, bool onlyInternal)
    {
        if (!(sheet.DrawingPatriarch is HSSFShapeContainer shapeContainer))
            return;
        foreach (var shape in shapeContainer.Children)
        {
            if (shape is HSSFPicture picture && picture.Anchor is HSSFClientAnchor anchor)
            {
                if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2, anchor.Col1,
                        anchor.Col2, onlyInternal))
                    shapeContainer.RemoveShape(picture);
            }
        }
    }

    /// <summary>
    /// 移除图片
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="minRow">最小行索引</param>
    /// <param name="maxRow">最大行索引</param>
    /// <param name="minCol">最小列索引</param>
    /// <param name="maxCol">最大列索引</param>
    /// <param name="onlyInternal">是否仅在内部</param>
    private static void RemovePictures(XSSFSheet sheet, int? minRow, int? maxRow, int? minCol,
        int? maxCol, bool onlyInternal)
    {
        throw new NotImplementedException($"{typeof(XSSFSheet)}尚未实现ClearPictures()方法");
        //foreach (var documentPart in sheet.GetRelations())
        //{
        //    if (documentPart is XSSFDrawing drawing)
        //    {
        //        foreach (var shape in drawing.GetShapes())
        //        {
        //            var picture = (XSSFPicture)shape;
        //            var anchor = picture.GetPreferredSize();
        //            if (!IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2,
        //                anchor.Col1,
        //                anchor.Col2, onlyInternal))
        //                continue;
        //        }
        //    }
        //}
    }

    /// <summary>
    /// 移动工作表中所有的图片
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    /// <param name="moveRowCount">移动行数</param>
    /// <param name="moveColCount">移动列数</param>
    public static void
        MovePictures(this NPOI.SS.UserModel.ISheet sheet, int moveRowCount = 0, int moveColCount = 0) =>
        sheet.MovePictures(null, null, null, null, true, moveRowCount, moveColCount);

    /// <summary>
    /// 移动工作表中指定区域的图片
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    /// <param name="minRow">最小行索引</param>
    /// <param name="maxRow">最大行索引</param>
    /// <param name="minCol">最小列索引</param>
    /// <param name="maxCol">最大列索引</param>
    /// <param name="onlyInternal">是否仅在内部</param>
    /// <param name="moveRowCount">移动行数</param>
    /// <param name="moveColCount">移动列数</param>
    public static void MovePictures(this NPOI.SS.UserModel.ISheet sheet, int? minRow, int? maxRow, int? minCol,
        int? maxCol, bool onlyInternal = true, int moveRowCount = 0, int moveColCount = 0)
    {
        if (!(sheet.DrawingPatriarch is HSSFShapeContainer shapeContainer))
            return;
        foreach (var shape in shapeContainer.Children)
        {
            if (shape.Anchor is IClientAnchor anchor)
            {
                if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2, anchor.Col1,
                        anchor.Col2, onlyInternal))
                {
                    anchor.Row1 += moveRowCount;
                    anchor.Row2 += moveRowCount;
                    anchor.Col1 += moveColCount;
                    anchor.Col2 += moveColCount;
                }
            }
        }
    }

    /// <summary>
    /// 尝试添加图片到工作表当中
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="row">行索引</param>
    /// <param name="col">列索引</param>
    /// <param name="pictureData">图片数据</param>
    /// <returns>添加成功则返回true</returns>
    public static bool TryAddPicture(this ISheet sheet, int row, int col, IPictureData pictureData) => TryAddPicture(sheet, row, col, pictureData.Data, pictureData.PictureType);

    /// <summary>
    /// 尝试添加图片到工作表当中
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="row">行索引</param>
    /// <param name="col">列索引</param>
    /// <param name="pictureBytes">图片数据</param>
    /// <param name="pictureType">图片类型</param>
    /// <returns>添加成功则返回true</returns>
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