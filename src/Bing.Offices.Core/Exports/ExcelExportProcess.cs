using System;
using System.Collections.Generic;
using System.IO;
using Bing.Offices.Abstractions.Exports;
using Bing.Utils.Helpers;
using Bing.Utils.IO;

namespace Bing.Offices.Exports
{
    /// <summary>
    /// Excel 导出处理
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    public class ExcelExportProcess<TEntity> : IExcelExportProcess where TEntity : class, new()
    {
        /// <summary>
        /// 导出数据函数
        /// </summary>
        private readonly GetExportDataEvent<TEntity> _exportDataFunc;

        /// <summary>
        /// 导出配置
        /// </summary>
        internal ExportConfig<TEntity> Config { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ExcelExportProcess{TEntity}"/>类型的实例
        /// </summary>
        /// <param name="func">导出数据函数</param>
        /// <param name="condition">查询条件</param>
        public ExcelExportProcess(GetExportDataEvent<TEntity> func, object condition)
        {
            this._exportDataFunc = func ?? throw new ArgumentNullException($"没有找到导出数据的方法!", nameof(func));
            this.Config = new ExportConfig<TEntity>(condition);
        }

        /// <summary>
        /// 执行方法
        /// </summary>
        /// <param name="baseDir">基础路径</param>
        /// <param name="fileName">文件名。不包含扩展名</param>
        public string Run(string baseDir, string fileName) => Run(baseDir, fileName, string.Empty);

        /// <summary>
        /// 执行方法
        /// </summary>
        /// <param name="baseDir">基础路径</param>
        /// <param name="fileName">文件名。不包含扩展名</param>
        /// <param name="exportFields">导出字段。以","分割</param>
        public string Run(string baseDir, string fileName, string exportFields)
        {
            var queryCount = DefaultSettings.DefaultExcelSetting.MaxExportNum;
            var data = this._exportDataFunc(this.Config.Condition, queryCount);
            try
            {
                return CreateFile(data, baseDir, fileName, exportFields);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 创建文件
        /// </summary>
        /// <param name="data">数据</param>
        /// <param name="baseDir">基础路径</param>
        /// <param name="fileName">文件名。不包含扩展名</param>
        /// <param name="exportFields">导出字段</param>
        private string CreateFile(IEnumerable<TEntity> data, string baseDir, string fileName, string exportFields)
        {
            baseDir = baseDir.TrimEnd('\\').TrimEnd('/');
            baseDir += "/";
            string absFilePath = Common.GetPhysicalPath(baseDir);
            DirectoryHelper.CreateIfNotExists(absFilePath);

            var newFileName = $"{fileName}_{DateTime.Now.ToString("yyyyMMddHHmmss_fff")}.xlsx";

            var filePath = Path.Combine(baseDir, newFileName);
            absFilePath = Path.Combine(absFilePath, newFileName);

            // TODO: 缺少生成文件操作
            return filePath;
        }
    }
}
