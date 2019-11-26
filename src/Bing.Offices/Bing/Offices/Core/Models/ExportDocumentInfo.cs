using System.Collections.Generic;
using System.Reflection;
using Bing.Offices.Core.Extension;

namespace Bing.Offices.Core.Models
{
    /// <summary>
    /// 导出文档信息
    /// </summary>
    /// <typeparam name="TData">数据类型</typeparam>
    public class ExportDocumentInfo<TData> where TData : class
    {
        /// <summary>
        /// 初始化一个<see cref="ExportDocumentInfo{TData}"/>类型的实例
        /// </summary>
        /// <param name="data">数据</param>
        public ExportDocumentInfo(TData data)
        {
            Headers = new List<ExportHeaderAttribute>();
            Data = data;
            Title = typeof(TData).GetCustomAttribute<ExportAttribute>()?.Name ?? typeof(TData).Name;
            foreach (var propertyInfo in typeof(TData).GetProperties())
            {
                var exportHeader = propertyInfo.PropertyType.GetCustomAttribute<ExportHeaderAttribute>() ??
                                   new ExportHeaderAttribute()
                                   {
                                       DisplayName = propertyInfo.GetDisplayName() ?? propertyInfo.Name
                                   };
                Headers.Add(exportHeader);
            }
        }

        /// <summary>
        /// 文档标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 头部信息
        /// </summary>
        public IList<ExportHeaderAttribute> Headers { get; set; }

        /// <summary>
        /// 数据
        /// </summary>
        public TData Data { get; set; }
    }

    /// <summary>
    /// 导出列表的文档信息
    /// </summary>
    /// <typeparam name="TData">数据类型</typeparam>
    public class ExportDocumentInfoOfListData<TData> where TData : class
    {
        /// <summary>
        /// 初始化一个<see cref="ExportDocumentInfoOfListData{TData}"/>类型的实例
        /// </summary>
        /// <param name="data">数据</param>
        public ExportDocumentInfoOfListData(ICollection<TData> data)
        {
            Headers = new List<ExportHeaderAttribute>();
            Data = data;
            Title = typeof(TData).GetCustomAttribute<ExportAttribute>()?.Name ?? typeof(TData).Name;
            foreach (var propertyInfo in typeof(TData).GetProperties())
            {
                var exportHeader = propertyInfo.PropertyType.GetCustomAttribute<ExportHeaderAttribute>() ??
                                   new ExportHeaderAttribute()
                                   {
                                       DisplayName = propertyInfo.GetDisplayName() ?? propertyInfo.Name
                                   };
                Headers.Add(exportHeader);
            }
        }

        /// <summary>
        /// 文档标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 头部信息
        /// </summary>
        public IList<ExportHeaderAttribute> Headers { get; set; }

        /// <summary>
        /// 数据
        /// </summary>
        public ICollection<TData> Data { get; set; }
    }
}
