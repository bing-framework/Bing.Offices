namespace Bing.Offices;

/// <summary>
/// 正则常量
/// </summary>
public static class RegexConst
{
    /// <summary>
    /// 邮箱
    /// </summary>
    public const string Email= @"\w[-\w.+]*@([A-Za-z0-9][-A-Za-z0-9]+\.)+[A-Za-z]{2,14}";

    /// <summary>
    /// 国内手机号
    /// </summary>
    public const string ChinaMobile= @"0?(13|14|15|17|18|19)[0-9]{9}";

    /// <summary>
    /// 身份证
    /// </summary>
    public const string IdCard= @"^(^\d{15}$|^\d{18}$|^\d{17}(\d|X|x))$";

    /// <summary>
    /// 非空
    /// </summary>
    public const string NotEmpty = @"\S";

    /// <summary>
    /// 性别
    /// </summary>
    public const string Gender= @"^男$|^女$|^先生$|^女士$|^Male|^Female$";

    /// <summary>
    /// 网址URL
    /// </summary>
    public const string Url = @"^((https|http|ftp|rtsp|mms)?:\/\/)[^\s]+";

    /// <summary>
    /// 国内车牌号
    /// </summary>
    public const string ChinaCarCode= @"^[京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领A-Z]{1}[A-Z]{1}[A-Z0-9]{4}[A-Z0-9挂学警港澳]{1}$";
}