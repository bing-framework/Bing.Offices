namespace Bing.Offices.Excel.Settings
{
    /// <summary>
    /// 验证设置
    /// </summary>
    internal class ValidateSetting
    {
        /// <summary>
        /// 是否允许为空。默认可空
        /// </summary>
        public bool IsNullable { get; set; } = true;

        /// <summary>
        /// 是否必填
        /// </summary>
        public bool IsRequired { get; set; }

        /// <summary>
        /// 最小长度
        /// </summary>
        public int? MinLength { get; set; }

        /// <summary>
        /// 最大长度
        /// </summary>
        public int? MaxLength { get; set; }

        /// <summary>
        /// 正则表达式字符串
        /// </summary>
        public string RegexString { get; set; }

        /// <summary>
        /// 是否允许重复数据
        /// </summary>
        public bool IsRepeat { get; set; }
    }
}
