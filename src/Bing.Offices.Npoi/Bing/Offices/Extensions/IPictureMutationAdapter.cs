using Bing.Offices.Exceptions;
using Bing.Offices.Metadata;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// 图片写入阶段适配器；测试可替换以确定性验证失败原子性合同。
/// </summary>
internal interface IPictureMutationAdapter
{
    /// <summary>
    /// 将图片数据注册到工作簿并返回图片索引。
    /// </summary>
    /// <param name="sheet">用于访问目标工作簿的工作表。</param>
    /// <param name="pictureBytes">图片二进制内容。</param>
    /// <param name="pictureType">图片格式。</param>
    /// <returns>工作簿中新增图片的索引。</returns>
    int AddPicture(ISheet sheet, byte[] pictureBytes, PictureType pictureType);

    /// <summary>
    /// 创建工作簿图片使用的客户端锚点。
    /// </summary>
    /// <param name="sheet">用于访问目标工作簿的工作表。</param>
    /// <returns>新建的客户端锚点。</returns>
    IClientAnchor CreateClientAnchor(ISheet sheet);

    /// <summary>
    /// 获取工作表现有的绘图容器，必要时创建新的容器。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <returns>工作表的绘图容器。</returns>
    IDrawing GetOrCreateDrawing(ISheet sheet);

    /// <summary>
    /// 根据锚点和图片索引创建图片形状。
    /// </summary>
    /// <param name="drawing">目标绘图容器。</param>
    /// <param name="anchor">图片位置锚点。</param>
    /// <param name="pictureIndex">工作簿中的图片索引。</param>
    /// <returns>新建的图片形状。</returns>
    IPicture CreatePicture(IDrawing drawing, IClientAnchor anchor, int pictureIndex);

    /// <summary>
    /// 按图片原始尺寸调整图片形状大小。
    /// </summary>
    /// <param name="picture">待调整的图片形状。</param>
    void Resize(IPicture picture);
}
