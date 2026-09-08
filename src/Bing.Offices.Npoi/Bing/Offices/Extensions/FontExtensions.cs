using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// 字体属性配置扩展，返回同一字体以支持链式调用。
/// </summary>
public static class FontExtensions
{
    #region SetFontHeightInPoints(设置字体大小)

    /// <summary>
    /// 设置字体磅值并返回当前字体。
    /// </summary>
    /// <param name="font">字体</param>
    /// <param name="fontSize">字体大小</param>
    /// <returns>当前字体。</returns>
    public static IFont SetFontHeightInPoints(this IFont font, short fontSize)
    {
        font.FontHeightInPoints = fontSize;
        return font;
    }

    #endregion

    #region SetColor(设置字体颜色)

    /// <summary>
    /// 设置字体索引颜色并返回当前字体。
    /// </summary>
    /// <param name="font">字体</param>
    /// <param name="color">颜色</param>
    /// <returns>当前字体。</returns>
    public static IFont SetColor(this IFont font, short color)
    {
        font.Color = color;
        return font;
    }

    #endregion

    #region SetBoldWeight(设置粗体)

    /// <summary>
    /// 根据粗体权重设置字体是否加粗并返回当前字体。
    /// </summary>
    /// <param name="font">字体</param>
    /// <param name="boldWeight">粗体大小</param>
    /// <returns>当前字体。</returns>
    public static IFont SetBoldWeight(this IFont font, short boldWeight)
    {
        font.IsBold = boldWeight >= 700;
        return font;
    }

    #endregion

    #region DefaultFont(默认字体)

    /// <summary>
    /// 将字体设置为库默认的宋体九磅样式并返回当前字体。
    /// </summary>
    /// <param name="font">字体</param>
    /// <returns>当前字体。</returns>
    public static IFont DefaultFont(this IFont font)
    {
        font.FontName = "宋体";
        font.FontHeightInPoints = 9;
        return font;
    }

    #endregion
}
