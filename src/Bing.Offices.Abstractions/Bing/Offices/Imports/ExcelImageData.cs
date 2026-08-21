using System;

namespace Bing.Offices.Imports;

/// <summary>
/// 与 Excel 提供程序无关的图片数据及其左上角锚点。
/// </summary>
public sealed class ExcelImageData
{
    /// <summary>
    /// 创建图片数据。
    /// </summary>
    public ExcelImageData(byte[] bytes, string contentType, int rowNumber, int columnNumber,
        int endRowNumber = 0, int endColumnNumber = 0)
    {
        if (bytes == null || bytes.Length == 0)
            throw new ArgumentException("图片数据不能为空。", nameof(bytes));
        Bytes = bytes;
        ContentType = contentType;
        RowNumber = rowNumber;
        ColumnNumber = columnNumber;
        EndRowNumber = endRowNumber;
        EndColumnNumber = endColumnNumber;
    }

    /// <summary>获取图片字节。</summary>
    public byte[] Bytes { get; }

    /// <summary>获取 MIME 类型。</summary>
    public string ContentType { get; }

    /// <summary>获取一开始的锚点行号。</summary>
    public int RowNumber { get; }

    /// <summary>获取一开始的锚点列号。</summary>
    public int ColumnNumber { get; }

    /// <summary>获取结束锚点行号。</summary>
    public int EndRowNumber { get; }

    /// <summary>获取结束锚点列号。</summary>
    public int EndColumnNumber { get; }
}
