using System;

namespace Bing.Offices.Exceptions
{
    /// <summary>
    /// Office空行异常
    /// </summary>
    [Serializable]
    public class OfficeEmptyLineException : OfficeException
    {
        /// <summary>
        /// 初始化一个<see cref="OfficeEmptyLineException"/>类型的实例
        /// </summary>
        /// <param name="message">序列化信息</param>
        public OfficeEmptyLineException(string message) : base(message) { }

        /// <summary>
        /// 初始化一个<see cref="OfficeEmptyLineException"/>类型的实例
        /// </summary>
        /// <param name="message">序列化信息</param>
        /// <param name="innerException">错误来源</param>
        public OfficeEmptyLineException(string message, Exception innerException) : base(message, innerException) { }
    }
}
