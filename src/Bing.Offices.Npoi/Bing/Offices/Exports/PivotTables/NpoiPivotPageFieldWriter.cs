using System.Collections.ObjectModel;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Bing.Offices.Exceptions;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Exports.PivotTables;

/// <summary>
/// 补齐原生透视单选索引零的序列化属性。
/// </summary>
internal static class NpoiPivotPageFieldWriter
{
    /// <summary>
    /// 获取需要显式保存的单选索引零快照。
    /// </summary>
    /// <param name="workbook">已完成透视配置的工作簿。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>按真实透视部件地址分组的字段索引与选择索引。</returns>
    /// <remarks>
    /// NPOI 2.7.4 会省略值为零的 pageField.item，即使 itemSpecified 已设置。
    /// 未指定 item 的全选字段和多选字段不进入补齐集合。
    /// </remarks>
    internal static IReadOnlyDictionary<string, IReadOnlyDictionary<int, uint>> Capture(
        XSSFWorkbook workbook, CancellationToken token)
    {
        var result = new Dictionary<string, IReadOnlyDictionary<int, uint>>(StringComparer.Ordinal);
        for (var sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
        {
            token.ThrowIfCancellationRequested();
            foreach (var pivot in ((XSSFSheet)workbook.GetSheetAt(sheetIndex)).GetPivotTables())
            {
                token.ThrowIfCancellationRequested();
                var pages = pivot.GetCTPivotTableDefinition().pageFields?.pageField;
                if (pages == null)
                    continue;
                var selections = pages.Where(page => page.itemSpecified && page.item == 0)
                    .ToDictionary(page => page.fld, page => page.item);
                if (selections.Count != 0)
                    result.Add(pivot.GetPackagePart().PartName.Name.TrimStart('/'),
                        new ReadOnlyDictionary<int, uint>(selections));
            }
        }
        return new ReadOnlyDictionary<string, IReadOnlyDictionary<int, uint>>(result);
    }

    /// <summary>
    /// 补齐指定透视部件的单选属性。
    /// </summary>
    /// <param name="source">包含原生透视定义的流。</param>
    /// <param name="target">接收补齐后定义的流，所有权归调用方。</param>
    /// <param name="selections">来源字段索引与明确选择的项目索引。</param>
    /// <param name="token">取消令牌。</param>
    internal static void Write(Stream source, Stream target, IReadOnlyDictionary<int, uint> selections,
        CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        using var reader = XmlReader.Create(source, new XmlReaderSettings
        { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        var document = XDocument.Load(reader);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var pages = document.Root?.Element(ns + "pageFields")?.Elements(ns + "pageField").ToArray()
            ?? Array.Empty<XElement>();
        foreach (var selection in selections)
        {
            token.ThrowIfCancellationRequested();
            var page = pages.SingleOrDefault(field => (int?)field.Attribute("fld") == selection.Key)
                ?? throw new BingOfficesConfigurationException("原生透视缺少待保存单选项的筛选字段。",
                    stage: BingOfficesStage.Write);
            page.SetAttributeValue("item", selection.Value);
        }
        using var writer = XmlWriter.Create(target, new XmlWriterSettings
        { Encoding = new UTF8Encoding(false), CloseOutput = false });
        document.Save(writer);
        token.ThrowIfCancellationRequested();
    }
}
