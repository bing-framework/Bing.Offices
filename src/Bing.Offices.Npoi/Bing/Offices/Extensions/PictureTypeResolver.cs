using System.Globalization;
using System.Runtime.CompilerServices;
using Bing.Helpers;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// 根据图片文件签名解析 NPOI 图片类型。
/// </summary>
internal static class PictureTypeResolver
{
    /// <summary>
    /// 解析图片类型。
    /// </summary>
    /// <param name="pictureBytes">图片字节。</param>
    /// <returns>可识别类型；未知内容保持原有 PNG 回退。</returns>
    public static PictureType Resolve(byte[] pictureBytes)
    {
        if (pictureBytes?.Length >= 8
            && pictureBytes[0] == 0x89
            && pictureBytes[1] == 0x50
            && pictureBytes[2] == 0x4E
            && pictureBytes[3] == 0x47
            && pictureBytes[4] == 0x0D
            && pictureBytes[5] == 0x0A
            && pictureBytes[6] == 0x1A
            && pictureBytes[7] == 0x0A)
            return PictureType.PNG;
        if (pictureBytes?.Length >= 3
            && pictureBytes[0] == 0xFF
            && pictureBytes[1] == 0xD8
            && pictureBytes[2] == 0xFF)
            return PictureType.JPEG;
        if (pictureBytes?.Length >= 6
            && pictureBytes[0] == 0x47
            && pictureBytes[1] == 0x49
            && pictureBytes[2] == 0x46
            && pictureBytes[3] == 0x38
            && (pictureBytes[4] == 0x37 || pictureBytes[4] == 0x39)
            && pictureBytes[5] == 0x61)
            return PictureType.GIF;
        return PictureType.PNG;
    }
}
