using System;
using System.ComponentModel;

namespace Bing.Offices.Tests.Models
{
    /// <summary>
    /// 导入测试样例
    /// </summary>
    public class ImportSample
    {
        /// <summary>
        /// 测试值
        /// </summary>
        public string TestValue { get; set; }

        /// <summary>
        /// 电子邮件验证
        /// </summary>
        public string Email { get; set; }

        /// <summary>
        /// 网址验证
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// 最大长度
        /// </summary>
        public string MaxLengthValue { get; set; }

        /// <summary>
        /// decimal值
        /// </summary>
        public decimal DecimalValue { get; set; }

        /// <summary>
        /// 可空decimal值
        /// </summary>
        public decimal? NullableDecimalValue { get; set; }

        /// <summary>
        /// float值
        /// </summary>
        public float FloatValue { get; set; }

        /// <summary>
        /// 可空float值
        /// </summary>
        public float? NullableFloatValue { get; set; }

        /// <summary>
        /// double值
        /// </summary>
        public double DoubleValue { get; set; }

        /// <summary>
        /// 可空double值
        /// </summary>
        public double? NullableDoubleValue { get; set; }

        /// <summary>
        /// bool值
        /// </summary>
        public bool BoolValue { get; set; }

        /// <summary>
        /// 可空bool值
        /// </summary>
        public bool? NullableBoolValue { get; set; }

        /// <summary>
        /// DateTime值
        /// </summary>
        public DateTime DateValue { get; set; }

        /// <summary>
        /// 可空DateTime值
        /// </summary>
        public DateTime? NullableDateValue { get; set; }

        /// <summary>
        /// 不可空枚举值
        /// </summary>
        public Gender LogLevel { get; set; }

        /// <summary>
        /// 可空枚举值
        /// </summary>
        public Gender? NullableLogLevel { get; set; }

        /// <summary>
        /// int值
        /// </summary>
        [Description("IntValue")]
        public int IntValue { get; set; }

        /// <summary>
        /// 可空int值
        /// </summary>
        public int? NullableIntValue { get; set; }

        /// <summary>
        /// short值
        /// </summary>
        public short ShortValue { get; set; }

        /// <summary>
        /// 可空short值
        /// </summary>
        public short? NullableShortValue { get; set; }

        /// <summary>
        /// long值
        /// </summary>
        public long LongValue { get; set; }

        /// <summary>
        /// 可空long值
        /// </summary>
        public long? NullableLongValue { get; set; }

        /// <summary>
        /// 忽略值
        /// </summary>
        public string IgnoreValue { get; set; }
    }
}
