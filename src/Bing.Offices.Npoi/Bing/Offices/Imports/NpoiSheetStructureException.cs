using System.Reflection;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Metadata;
using Bing.Offices.Providers;
using NPOI.SS.UserModel;

namespace Bing.Offices.Imports;

/// <summary>
/// 表示导入工作表结构不符合请求的异常。
/// </summary>
internal sealed class NpoiSheetStructureException : InvalidOperationException
{
    /// <summary>
    /// 初始化一个 <see cref="NpoiSheetStructureException" /> 类型的实例。
    /// </summary>
    /// <param name="message">描述工作表结构错误的消息。</param>
    internal NpoiSheetStructureException(string message) : base(message) { }
}
