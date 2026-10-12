using System.Globalization;
using System.Text;
using System.Xml;

namespace Bing.Offices.Npoi.Exports.PivotTables;

/// <summary>
/// 写入完整的 OOXML 透视缓存记录。
/// </summary>
/// <remarks>
/// NPOI 2.7.4 的工作簿写入未保存公开缓存对象中的记录，因而显式写入已验证的共享项索引。
/// 缓存定义和包关系继续由 NPOI 创建；此写入器不计算或替换透视结果。
/// </remarks>
internal static class NpoiPivotCacheRecordWriter
{
    /// <summary>
    /// SpreadsheetML 缓存记录使用的命名空间。
    /// </summary>
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// 将完整缓存记录写入目标流。
    /// </summary>
    /// <param name="destination">接收记录 XML 的可写流，所有权归调用方。</param>
    /// <param name="records">已经隔离并校验的缓存记录快照。</param>
    /// <param name="token">取消令牌。</param>
    internal static void Write(Stream destination, NpoiPivotCacheRecords records, CancellationToken token)
    {
        if (destination == null)
            throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite)
            throw new ArgumentException("目标流不可写入。", nameof(destination));
        if (records == null)
            throw new ArgumentNullException(nameof(records));
        token.ThrowIfCancellationRequested();
        using var writer = XmlWriter.Create(destination, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            CloseOutput = false,
            Indent = false
        });
        writer.WriteStartDocument();
        writer.WriteStartElement("pivotCacheRecords", SpreadsheetNamespace);
        writer.WriteAttributeString("count", records.Rows.Count.ToString(CultureInfo.InvariantCulture));
        foreach (var row in records.Rows)
        {
            token.ThrowIfCancellationRequested();
            writer.WriteStartElement("r", SpreadsheetNamespace);
            foreach (var index in row)
            {
                token.ThrowIfCancellationRequested();
                writer.WriteStartElement("x", SpreadsheetNamespace);
                writer.WriteAttributeString("v", index.ToString(CultureInfo.InvariantCulture));
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
        writer.WriteEndDocument();
        writer.Flush();
        token.ThrowIfCancellationRequested();
    }
}
