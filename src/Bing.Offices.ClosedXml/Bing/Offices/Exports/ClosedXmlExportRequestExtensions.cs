using System.Collections;
using System.Globalization;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Offices.ClosedXml.Entities;
using Bing.Offices.ClosedXml.Internals;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.IO;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Exports;

/// <summary>
/// 提供 ClosedXML 导出请求的 Provider 辅助访问器。
/// </summary>
internal static class ClosedXmlExportRequestExtensions
{
    /// <summary>
    /// 汇总 Workbook 及其各 Sheet 的图表定义。
    /// </summary>
    /// <param name="request">Workbook 导出请求。</param>
    /// <returns>图表定义集合。</returns>
    public static IReadOnlyList<ExcelChartDefinition> Charts(this ExcelWorkbookExportRequest request) =>
        request.Sheets.SelectMany(sheet => sheet.Charts ?? Array.Empty<ExcelChartDefinition>()).ToArray();
}
