using System.Collections.Generic;

namespace Bing.Offices.Abstractions.Imports
{
    /// <summary>
    /// 导入选项配置
    /// </summary>
    public interface IImportOptions
    {
        /// <summary>
        /// 文件地址
        /// </summary>
        string FileUrl { get; set; }

        /// <summary>
        /// 工作表索引
        /// </summary>
        int SheetIndex { get; set; }

        /// <summary>
        /// 表头行索引
        /// </summary>
        int HeaderRowIndex { get; set; }

        /// <summary>
        /// 数据行索引
        /// </summary>
        int DataRowIndex { get; set; }

        /// <summary>
        /// 映射字典
        /// </summary>
        IDictionary<string,string> MappingDictionary { get; set; }

        /// <summary>
        /// 校验模式
        /// </summary>
        ValidateMode ValidateMode { get; set; }

        /// <summary>
        /// 自定义导入提供程序
        /// </summary>
        IExcelImportProvider CustomImportProvider { get; set; }
    }
}
