using Bing.Offices.Enums.Styles;

namespace Bing.Offices.Npoi.Resolvers;

/// <summary>
/// 对齐方式解析器
/// </summary>
public class AlignmentResolver
{
    /// <summary>
    /// 解析
    /// </summary>
    /// <param name="alignment">水平对齐方式</param>
    public static NPOI.SS.UserModel.HorizontalAlignment Resolve(Bing.Offices.Enums.Styles.HorizontalAlignment alignment)
    {
        switch (alignment)
        {
            case HorizontalAlignment.Left:
                return NPOI.SS.UserModel.HorizontalAlignment.Left;
            case HorizontalAlignment.Center:
                return NPOI.SS.UserModel.HorizontalAlignment.Center;
            case HorizontalAlignment.Right:
                return NPOI.SS.UserModel.HorizontalAlignment.Right;
            default:
                return NPOI.SS.UserModel.HorizontalAlignment.General;
        }
    }

    /// <summary>
    /// 解析
    /// </summary>
    /// <param name="alignment">垂直对齐方式</param>
    public static NPOI.SS.UserModel.VerticalAlignment Resolve(Bing.Offices.Enums.Styles.VerticalAlignment alignment)
    {
        switch (alignment)
        {
            case VerticalAlignment.Top:
                return NPOI.SS.UserModel.VerticalAlignment.Top;
            case VerticalAlignment.Center:
                return NPOI.SS.UserModel.VerticalAlignment.Center;
            case VerticalAlignment.Bottom:
                return NPOI.SS.UserModel.VerticalAlignment.Bottom;
            default:
                return NPOI.SS.UserModel.VerticalAlignment.None;
        }
    }
}
