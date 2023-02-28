using System;
using System.Linq;
using System.Collections.Generic;

namespace Bing.Offices.Imports;

/// <summary>
/// 导入结果
/// </summary>
public class ImportResult
{
    /// <summary>
    /// 初始化一个<see cref="ImportResult"/>类型的实例
    /// </summary>
    public ImportResult()
    {
        RowErrors = new List<DataRowErrorInfo>();
    }

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
                                    (TemplateErrors?.Count(p => p.ErrorLevel == ErrorLevels.Error) ?? 0) > 0 ||
                                    (RowErrors?.Count ?? 0) > 0;
}

/// <summary>
/// 导入结果
/// </summary>
/// <typeparam name="T">对象类型</typeparam>
public class ImportResult<T> : ImportResult
    where T : class
{
    /// <summary>
    /// 导入数据
    /// </summary>
    public virtual ICollection<T> Data { get; set; }
}