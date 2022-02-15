using System;
using System.Data;
using System.IO;
using System.Text;
using Bing.Offices.Internals;
using Bing.Text;

namespace Bing.Offices
{
    /// <summary>
    /// Csv帮助类
    /// </summary>
    public static class CsvHelper
    {
        /// <summary>
        /// Csv 分隔符字符。默认：','
        /// </summary>
        public static char CsvSeparatorCharacter { get; set; } = ',';

        /// <summary>
        /// Csv 引用字符。默认：'"'
        /// </summary>
        public static char CsvQuoteCharacter = '"';

        /// <summary>
        /// 转换为Csv文件
        /// </summary>
        /// <param name="dataTable">数据表</param>
        /// <param name="filePath">文件路径</param>
        public static bool ToCsvFile(DataTable dataTable, string filePath) => ToCsvFile(dataTable, filePath, true);

        /// <summary>
        /// 转换为Csv文件
        /// </summary>
        /// <param name="dataTable">数据表</param>
        /// <param name="filePath">文件路径</param>
        /// <param name="includeHeader">是否包含表头</param>
        public static bool ToCsvFile(DataTable dataTable, string filePath, bool includeHeader)
        {
            if (dataTable == null)
                throw new ArgumentNullException(nameof(dataTable));
            var dir = Path.GetDirectoryName(filePath);
            if (dir == null)
                throw new ArgumentException("无效文件路径", nameof(filePath));
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            var csvText = GetCsvText(dataTable, includeHeader);
            if (string.IsNullOrWhiteSpace(csvText))
                return false;
            File.WriteAllText(filePath, csvText, Encoding.UTF8);
            return true;
        }

        /// <summary>
        /// 转换为Csv字节数组
        /// </summary>
        /// <param name="dataTable">数据表</param>
        public static byte[] ToCsvBytes(DataTable dataTable) => ToCsvBytes(dataTable, true);

        /// <summary>
        /// 转换为Csv字节数组
        /// </summary>
        /// <param name="dataTable">数据表</param>
        /// <param name="includeHeader">是否包含表头</param>
        public static byte[] ToCsvBytes(DataTable dataTable, bool includeHeader) => GetCsvText(dataTable, includeHeader).ToBytes();

        /// <summary>
        /// 获取Csv文本
        /// </summary>
        /// <param name="dataTable">数据表</param>
        /// <param name="includeHeader">是否包含表头</param>
        public static string GetCsvText(DataTable dataTable, bool includeHeader = true)
        {
            if (dataTable == null || dataTable.Rows.Count == 0 || dataTable.Columns.Count == 0)
                return string.Empty;
            var result = new StringBuilder();
            if (includeHeader)
            {
                for (var i = 0; i < dataTable.Columns.Count; i++)
                {
                    if (i > 0)
                        result.Append(CsvSeparatorCharacter);
                    var columnName = InternalHelper.GetDecodeColumnName(dataTable.Columns[i].ColumnName);
                    result.Append(columnName);
                }
                result.AppendLine();
            }

            for (var i = 0; i < dataTable.Rows.Count; i++)
            {
                for (var j = 0; j < dataTable.Columns.Count; j++)
                {
                    if (j > 0)
                        result.Append(CsvSeparatorCharacter);
                    // https://stackoverflow.com/questions/4617935/is-there-a-way-to-include-commas-in-csv-columns-without-breaking-the-formatting
                    var val = dataTable.Rows[i][j]?.ToString()?.Replace("\"", "\"\"");
                    if (val != null && val.Length > 0)
                        result.Append(val.IndexOf(CsvSeparatorCharacter) > -1 ? $"\"{val}\"" : val);
                }
                result.AppendLine();
            }
            return result.ToString();
        }
    }
}
