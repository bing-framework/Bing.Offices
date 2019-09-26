using System;
using System.ComponentModel;
using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models
{
    /// <summary>
    /// 导入用户
    /// </summary>
    public class ImportUser
    {
        /// <summary>
        /// 标识
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 车牌号
        /// </summary>
        [ColumnName("车牌号")]
        [Required]
        [Regex("^[京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领A-Z]{1}[A-Z]{1}[A-Z0-9]{4}[A-Z0-9挂学警港澳]{1}$")]
        [Duplication]
        public string CarCode { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        [ColumnName("姓名")]
        [MaxLength(10)]
        public string Name { get; set; }

        /// <summary>
        /// 身份证
        /// </summary>
        [ColumnName("身份证")]
        [Regex(@"^(^\d{15}$|^\d{18}$|^\d{17}(\d|X|x))$")]
        public string IdCard { get; set; }

        /// <summary>
        /// 手机号
        /// </summary>
        [ColumnName("手机号")]
        [Regex("0?(13|14|15|17|18|19)[0-9]{9}")]
        public string Phone { get; set; }

        /// <summary>
        /// 年龄
        /// </summary>
        [ColumnName("年龄")]
        [Range(0,100)]
        public int? Age { get; set; }

        /// <summary>
        /// 性别
        /// </summary>
        [ColumnName("性别")]
        [Regex("^先生$|^女士$")]
        public Gender Gender { get; set; }

        /// <summary>
        /// 注册日期
        /// </summary>
        [ColumnName("注册日期")]
        [DateTime]
        public DateTime RegisterTime { get; set; }
    }

    /// <summary>
    /// 性别
    /// </summary>
    public enum Gender
    {
        /// <summary>
        /// 女
        /// </summary>
        [Description("女士")]
        Female = 1,
        /// <summary>
        /// 男
        /// </summary>
        [Description("先生")]
        Male = 2
    }
}
