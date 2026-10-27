namespace Bing.Offices.Exports;

/// <summary>使用有限行窗口生成 XLSX；不支持模板、图片、图表或合并区域。</summary>
public interface IExcelStreamingExporter
{
    /// <summary>生成完整 XLSX 后复制到调用方拥有的流，保持流打开。</summary>
    void Export(ExcelWorkbookExportRequest request, Stream destination,
        ExcelStreamingExportOptions options = null, CancellationToken cancellationToken = default);
    /// <summary>生成阶段同步，向调用方流复制使用真实异步 IO；保持流打开。</summary>
    Task ExportAsync(ExcelWorkbookExportRequest request, Stream destination,
        ExcelStreamingExportOptions options = null, CancellationToken cancellationToken = default);
    /// <summary>原子提交到文件；失败或取消不替换既有目标。</summary>
    void ExportToFile(ExcelWorkbookExportRequest request, string path,
        ExcelStreamingExportOptions options = null, CancellationToken cancellationToken = default);
    /// <summary>通过异步文件 IO 原子提交；SXSSF 构建和打包仍为同步阶段。</summary>
    Task ExportToFileAsync(ExcelWorkbookExportRequest request, string path,
        ExcelStreamingExportOptions options = null, CancellationToken cancellationToken = default);
}

/// <summary>低内存导出的有限资源额度；不代表进程堆或 NPOI 系统临时目录的硬配额。</summary>
public sealed class ExcelStreamingExportOptions
{
    /// <summary>驻留行数上限，默认 100。</summary>
    public int RowWindowSize { get; init; } = 100;
    /// <summary>是否压缩 NPOI 工作表临时 XML，默认启用。</summary>
    public bool CompressTemporaryFiles { get; init; } = true;
    /// <summary>列数上限，默认 100。</summary>
    public int MaxColumns { get; init; } = 100;
    /// <summary>Sheet 数量上限，默认 32。</summary>
    public int MaxSheets { get; init; } = 32;
    /// <summary>工作簿样式数量上限，默认 4096。</summary>
    public int MaxStyles { get; init; } = 4096;
    /// <summary>工作簿字体数量上限，默认 256。</summary>
    public int MaxFonts { get; init; } = 256;
    /// <summary>单元格文本字符上限，不能大于 XLSX 的 32767。</summary>
    public int MaxCellCharacters { get; init; } = 32767;
    /// <summary>单行文本字符上限，默认 262144。</summary>
    public int MaxRowCharacters { get; init; } = 262144;
    /// <summary>窗口内文本字符上限，默认 2097152。</summary>
    public int MaxBufferedCharacters { get; init; } = 2097152;
    /// <summary>库自有临时文件总额度，默认 512 MiB，不包括 NPOI 系统临时文件。</summary>
    public long MaxTemporaryBytes { get; init; } = 512L * 1024 * 1024;
    /// <summary>库自有临时文件父目录；null 使用系统临时目录。不改变 NPOI 的目录。</summary>
    public string TempDirectory { get; init; }
    /// <summary>验证正数额度及 Excel 格式边界。</summary>
    public void Validate()
    {
        if (RowWindowSize < 1 || MaxColumns < 1 || MaxColumns > 16384 || MaxSheets < 1
            || MaxStyles < 1 || MaxFonts < 1 || MaxCellCharacters < 1 || MaxCellCharacters > 32767
            || MaxRowCharacters < 1 || MaxBufferedCharacters < MaxRowCharacters || MaxTemporaryBytes < 1)
            throw new ArgumentOutOfRangeException(nameof(ExcelStreamingExportOptions), "流式导出额度必须为正数且符合 Excel 边界。");
    }
}
