using System.Data;
using System.Globalization;
using System.Text;
using Bing.Offices.Csv;
using Bing.Offices.Internals;

namespace Bing.Offices;

/// <summary>
/// DataTable CSV 兼容帮助类。
/// </summary>
public static class CsvHelper
{
    /// <summary>
    /// 旧版 CSV 分隔符；该可变全局状态仅用于兼容旧调用方。
    /// </summary>
    [Obsolete("请使用包含 delimiter 参数的 CSV 导出重载，避免跨请求共享可变状态。")]
    public static char CsvSeparatorCharacter { get; set; } = ',';

    /// <summary>
    /// 旧版 CSV 引用字符；该可变全局状态仅用于兼容旧调用方。
    /// </summary>
    [Obsolete("请使用包含 quote 参数的 CSV 导出重载，避免跨请求共享可变状态。")]
    public static char CsvQuoteCharacter = '"';

    /// <summary>
    /// 使用旧版全局分隔符和引用字符将 DataTable 写入 CSV 文件。
    /// </summary>
    /// <param name="dataTable">数据表</param>
    /// <param name="filePath">文件路径</param>
    /// <returns>文件成功写入时为 true。</returns>
    [Obsolete("请使用包含 delimiter 和 quote 参数的重载。")]
    public static bool ToCsvFile(DataTable dataTable, string filePath) => ToCsvFile(dataTable, filePath, true,
        CsvSeparatorCharacter, CsvQuoteCharacter);

    /// <summary>
    /// 使用旧版全局分隔符和引用字符将 DataTable 写入 CSV 文件。
    /// </summary>
    /// <param name="dataTable">数据表</param>
    /// <param name="filePath">文件路径</param>
    /// <param name="includeHeader">是否包含表头</param>
    /// <returns>文件成功写入时为 true。</returns>
    [Obsolete("请使用包含 delimiter 和 quote 参数的重载。")]
    public static bool ToCsvFile(DataTable dataTable, string filePath, bool includeHeader) =>
        ToCsvFile(dataTable, filePath, includeHeader, CsvSeparatorCharacter, CsvQuoteCharacter);

    /// <summary>
    /// 使用显式 CSV 格式选项将 DataTable 写入文件。
    /// </summary>
    /// <param name="dataTable">数据表。</param>
    /// <param name="filePath">文件路径。</param>
    /// <param name="includeHeader">是否包含表头。</param>
    /// <param name="delimiter">字段分隔符。</param>
    /// <param name="quote">字段引用字符。</param>
    /// <param name="formulaInjectionPolicy">潜在公式字段的处理策略。</param>
    /// <returns>文件成功写入时为 true。</returns>
    public static bool ToCsvFile(DataTable dataTable, string filePath, bool includeHeader, char delimiter = ',',
        char quote = '"', CsvFormulaInjectionPolicy formulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape)
    {
        if (dataTable == null)
            throw new ArgumentNullException(nameof(dataTable));
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("文件路径不能为空。", nameof(filePath));
        var dir = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);
        var csvText = GetCsvText(dataTable, includeHeader, delimiter, quote, formulaInjectionPolicy);
        File.WriteAllText(filePath, csvText, new UTF8Encoding(false));
        return true;
    }

    /// <summary>
    /// 使用旧版全局分隔符和引用字符将 DataTable 转换为 UTF-8 字节数组。
    /// </summary>
    /// <param name="dataTable">数据表</param>
    /// <returns>不带 UTF-8 BOM 的 CSV 字节数组。</returns>
    [Obsolete("请使用包含 delimiter 和 quote 参数的重载。")]
    public static byte[] ToCsvBytes(DataTable dataTable) => ToCsvBytes(dataTable, true, CsvSeparatorCharacter,
        CsvQuoteCharacter);

    /// <summary>
    /// 使用旧版全局分隔符和引用字符将 DataTable 转换为 UTF-8 字节数组。
    /// </summary>
    /// <param name="dataTable">数据表</param>
    /// <param name="includeHeader">是否包含表头</param>
    /// <returns>不带 UTF-8 BOM 的 CSV 字节数组。</returns>
    [Obsolete("请使用包含 delimiter 和 quote 参数的重载。")]
    public static byte[] ToCsvBytes(DataTable dataTable, bool includeHeader) => ToCsvBytes(dataTable, includeHeader,
        CsvSeparatorCharacter, CsvQuoteCharacter);

    /// <summary>
    /// 使用显式 CSV 格式选项将 DataTable 转换为 UTF-8 字节数组。
    /// </summary>
    /// <param name="dataTable">数据表。</param>
    /// <param name="includeHeader">是否包含表头。</param>
    /// <param name="delimiter">字段分隔符。</param>
    /// <param name="quote">字段引用字符。</param>
    /// <param name="formulaInjectionPolicy">潜在公式字段的处理策略。</param>
    /// <returns>不带 UTF-8 BOM 的 CSV 字节数组。</returns>
    public static byte[] ToCsvBytes(DataTable dataTable, bool includeHeader, char delimiter = ',', char quote = '"',
        CsvFormulaInjectionPolicy formulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape) =>
        new UTF8Encoding(false).GetBytes(GetCsvText(dataTable, includeHeader, delimiter, quote, formulaInjectionPolicy));

    /// <summary>
    /// 使用旧版全局分隔符和引用字符获取 DataTable 的 CSV 文本。
    /// </summary>
    /// <param name="dataTable">数据表</param>
    /// <param name="includeHeader">是否包含表头</param>
    /// <returns>CSV 文本；没有列时返回空字符串。</returns>
    [Obsolete("请使用包含 delimiter 和 quote 参数的重载。")]
    public static string GetCsvText(DataTable dataTable, bool includeHeader = true) => GetCsvText(dataTable, includeHeader,
        CsvSeparatorCharacter, CsvQuoteCharacter);

    /// <summary>
    /// 使用显式 CSV 格式选项将 DataTable 转换为文本。
    /// </summary>
    /// <param name="dataTable">数据表。</param>
    /// <param name="includeHeader">是否包含表头。</param>
    /// <param name="delimiter">字段分隔符。</param>
    /// <param name="quote">字段引用字符。</param>
    /// <param name="formulaInjectionPolicy">潜在公式字段的处理策略。</param>
    /// <returns>CSV 文本；没有列时返回空字符串。</returns>
    public static string GetCsvText(DataTable dataTable, bool includeHeader, char delimiter = ',', char quote = '"',
        CsvFormulaInjectionPolicy formulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape)
    {
        if (dataTable == null)
            throw new ArgumentNullException(nameof(dataTable));
        ValidateOptions(delimiter, quote, formulaInjectionPolicy);
        if (dataTable.Columns.Count == 0)
            return string.Empty;
        var result = new StringBuilder();
        using var writer = new StringWriter(result);
        if (includeHeader)
            CsvRecordWriter.Write(writer, dataTable.Columns.Cast<DataColumn>().Select(column =>
                InternalHelper.GetDecodeColumnName(column.ColumnName)), delimiter, quote, "\r\n", formulaInjectionPolicy);
        foreach (DataRow row in dataTable.Rows)
            CsvRecordWriter.Write(writer, dataTable.Columns.Cast<DataColumn>().Select(column =>
                Convert.ToString(row[column], CultureInfo.InvariantCulture) ?? string.Empty), delimiter, quote, "\r\n",
                formulaInjectionPolicy);
        return result.ToString();
    }

    /// <summary>验证 CSV 分隔符、引用字符和公式防护策略。</summary>
    /// <param name="delimiter">字段分隔符。</param>
    /// <param name="quote">字段引用字符。</param>
    /// <param name="formulaInjectionPolicy">潜在公式字段的处理策略。</param>
    private static void ValidateOptions(char delimiter, char quote, CsvFormulaInjectionPolicy formulaInjectionPolicy)
    {
        if (delimiter == quote || delimiter == '\r' || delimiter == '\n')
            throw new ArgumentOutOfRangeException(nameof(delimiter));
        if (quote == '\r' || quote == '\n')
            throw new ArgumentOutOfRangeException(nameof(quote));
        if (!Enum.IsDefined(typeof(CsvFormulaInjectionPolicy), formulaInjectionPolicy))
            throw new ArgumentOutOfRangeException(nameof(formulaInjectionPolicy));
    }
}
