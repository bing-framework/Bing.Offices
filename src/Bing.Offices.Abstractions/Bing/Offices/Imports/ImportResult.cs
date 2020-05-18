using System;
using System.Collections.Generic;
using System.Linq;

namespace Bing.Offices.Imports
{
    /// <summary>
    /// 导入结果
    /// </summary>
    /// <typeparam name="T">结果类型</typeparam>
    public class ImportResult<T> where T : class
    {
        /// <summary>
        /// 导入数据
        /// </summary>
        public virtual ICollection<T> Data { get; set; }

        /// <summary>
        /// 验证错误
        /// </summary>
        public virtual IList<DataRowErrorInfo> RowErrors { get; set; }

        /// <summary>
        /// 模板错误
        /// </summary>
        public virtual IList<TemplateErrorInfo> TemplateErrors { get; set; }

        /// <summary>
        /// 导入异常信息
        /// </summary>
        public virtual Exception Exception { get; set; }

        /// <summary>
        /// 是否存在导入错误
        /// </summary>
        public virtual bool HasError => Exception != null ||
                                        (TemplateErrors?.Count(x => x.ErrorLevel == ErrorLevel.Error) ?? 0) > 0 ||
                                        (RowErrors?.Count ?? 0) > 0;

        /// <summary>
        /// 导入的表头列表信息
        /// </summary>
        public virtual IList<ImportHeaderInfo> ImportHeaderInfos { get; set; }
    }
}
