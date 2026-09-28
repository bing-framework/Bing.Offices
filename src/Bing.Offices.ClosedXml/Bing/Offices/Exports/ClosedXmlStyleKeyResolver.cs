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
/// 将公共样式键解析为 ClosedXML Provider 支持的样式。
/// </summary>
internal static class ClosedXmlStyleKeyResolver
{
    /// <summary>
    /// 解析表头或正文样式键。
    /// </summary>
    /// <param name="key">样式键。</param>
    /// <param name="header">是否解析表头样式。</param>
    /// <returns>解析后的公共样式；键为空时返回 <see langword="null" />。</returns>
    internal static Styles.ExcelCellStyle Resolve(string key, bool header)
    {
        if (string.IsNullOrWhiteSpace(key))
            return null;
        if (string.Equals(key, "header", StringComparison.OrdinalIgnoreCase)
            || string.Equals(key, "bold", StringComparison.OrdinalIgnoreCase))
            return new Styles.ExcelCellStyle { Bold = true };
        if (string.Equals(key, "body", StringComparison.OrdinalIgnoreCase)
            || string.Equals(key, "default", StringComparison.OrdinalIgnoreCase))
            return new Styles.ExcelCellStyle();
        throw new BingOfficesConfigurationException($"未注册的{(header ? "表头" : "正文")}样式键: {key}",
            stage: BingOfficesStage.Plan);
    }
}
