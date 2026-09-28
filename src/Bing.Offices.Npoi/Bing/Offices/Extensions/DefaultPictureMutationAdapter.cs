using Bing.Offices.Exceptions;
using Bing.Offices.Metadata;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// 使用 NPOI API 执行图片写入和尺寸调整。
/// </summary>
internal sealed class DefaultPictureMutationAdapter : IPictureMutationAdapter
{
    /// <inheritdoc />
    public int AddPicture(ISheet sheet, byte[] pictureBytes, PictureType pictureType)
        => sheet.Workbook.AddPicture(pictureBytes, pictureType);

    /// <inheritdoc />
    public IClientAnchor CreateClientAnchor(ISheet sheet)
        => sheet.Workbook.GetCreationHelper().CreateClientAnchor();

    /// <inheritdoc />
    public IDrawing GetOrCreateDrawing(ISheet sheet)
        => sheet.DrawingPatriarch ?? sheet.CreateDrawingPatriarch();

    /// <inheritdoc />
    public IPicture CreatePicture(IDrawing drawing, IClientAnchor anchor, int pictureIndex)
        => drawing.CreatePicture(anchor, pictureIndex);

    /// <inheritdoc />
    public void Resize(IPicture picture) => picture.Resize();
}
