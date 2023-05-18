namespace Bing.Offices.Imports;

/// <summary>
/// Excel 导入器
/// </summary>
public class ExcelImporter : IExcelImporter
{
    /// <summary>
    /// 生成导入模板
    /// </summary>
    /// <typeparam name="T">类型</typeparam>
    /// <returns>二进制字节数组</returns>
    public async Task<byte[]> GenerateTemplateBytesAsync<T>() where T : class, new()
    {
        throw new NotImplementedException();
    }
}
