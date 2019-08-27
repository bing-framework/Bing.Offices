using System.IO;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Extensions
{
    /// <summary>
    /// 工作簿<see cref="IWorkbook"/> 扩展
    /// </summary>
    public static class WorkbookExtensions
    {
        /// <summary>
        /// 将工作簿转换成字节数组
        /// </summary>
        /// <param name="workbook">工作簿</param>
        public static byte[] SaveToBuffer(this IWorkbook workbook)
        {
            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                return ms.ToArray();
            }
        }

        /// <summary>
        /// 转换为工作簿
        /// </summary>
        /// <param name="workbookBytes">工作簿字节数组</param>
        public static IWorkbook ToWorkbook(this byte[] workbookBytes)
        {
            using (var stream = new MemoryStream(workbookBytes))
                return WorkbookFactory.Create(stream);
        }
    }
}
