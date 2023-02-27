using System;
using System.IO;
using Bing.Offices.Npoi.Internals;
using Bing.Offices.Settings;
using Bing.Text;
using NPOI.HPSF;

namespace Bing.Offices.Npoi;

/// <summary>
/// Excel 操作辅助类
/// </summary>
public static class ExcelHelper
{
    /// <summary>
    /// 应用程序版本
    /// </summary>
    private static readonly Version AppVersion = typeof(ExcelHelper).Assembly.GetName().Version!;

    /// <summary>
    /// 校验Excel文件路径
    /// </summary>
    /// <param name="excelPath">Excel文件路径</param>
    /// <param name="msg">错误消息</param>
    /// <param name="isExport">是否导出</param>
    /// <exception cref="ArgumentNullException"></exception>
    private static bool ValidateExcelFilePath(string excelPath, out string msg, bool isExport = false)
    {
        if (string.IsNullOrWhiteSpace(excelPath))
            throw new ArgumentNullException(nameof(excelPath));
        if (isExport || File.Exists(excelPath))
        {
            var ext = Path.GetExtension(excelPath);
            if (ext.EqualsIgnoreCase(".xls") || ext.EqualsIgnoreCase(".xlsx"))
            {
                msg = string.Empty;
                return true;
            }

            msg = "无效的Excel文件";
            return false;
        }

        msg = "找不到文件";
        return false;
    }

    #region PrepareWorkbook(准备工作簿)

    /// <summary>
    /// 准备工作簿
    /// </summary>
    /// <param name="excelPath">Excel文件路径</param>
    public static NPOI.SS.UserModel.IWorkbook PrepareWorkbook(string excelPath) => PrepareWorkbook(excelPath, null);

    /// <summary>
    /// 准备工作簿
    /// </summary>
    /// <param name="excelPath">Excel文件路径</param>
    /// <param name="excelSetting">Excel设置</param>
    /// <exception cref="ArgumentException"></exception>
    public static NPOI.SS.UserModel.IWorkbook PrepareWorkbook(string excelPath, ExcelSetting? excelSetting)
    {
        if (!ValidateExcelFilePath(excelPath, out var msg, true))
            throw new ArgumentException(msg);
        return PrepareWorkbook(!Path.GetExtension(excelPath).EqualsIgnoreCase(".xls"), excelSetting);
    }

    /// <summary>
    /// 准备工作簿
    /// </summary>
    /// <param name="format">Excel格式</param>
    /// <param name="excelSetting">Excel设置</param>
    public static NPOI.SS.UserModel.IWorkbook PrepareWorkbook(ExcelFormat format, ExcelSetting? excelSetting) =>
        PrepareWorkbook(format == ExcelFormat.Xlsx, excelSetting);

    /// <summary>
    /// 准备工作簿
    /// </summary>
    public static NPOI.SS.UserModel.IWorkbook PrepareWorkbook() => PrepareWorkbook(true);

    /// <summary>
    /// 准备工作簿
    /// </summary>
    /// <param name="format">Excel格式</param>
    public static NPOI.SS.UserModel.IWorkbook PrepareWorkbook(ExcelFormat format) => PrepareWorkbook(format == ExcelFormat.Xlsx);

    /// <summary>
    /// 准备工作簿
    /// </summary>
    /// <param name="isXlsx">是否Xlsx格式</param>
    public static NPOI.SS.UserModel.IWorkbook PrepareWorkbook(bool isXlsx) => PrepareWorkbook(isXlsx, null);

    /// <summary>
    /// 准备工作簿
    /// </summary>
    /// <param name="isXlsx">是否Xlsx格式</param>
    /// <param name="excelSetting">Excel设置</param>
    public static NPOI.SS.UserModel.IWorkbook PrepareWorkbook(bool isXlsx, ExcelSetting? excelSetting)
    {
        var setting = excelSetting ?? ExcelSetting.Default;
        if (isXlsx)
        {
            var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
            var props = workbook.GetProperties();
            props.CoreProperties.Creator = setting.Author;
            props.CoreProperties.Created = DateTime.Now;
            props.CoreProperties.Modified = DateTime.Now;
            props.CoreProperties.Title = setting.Title;
            props.CoreProperties.Subject = setting.Subject;
            props.CoreProperties.Category = setting.Category;
            props.CoreProperties.Description = setting.Description;
            props.ExtendedProperties.GetUnderlyingProperties().Company = setting.Company;
            props.ExtendedProperties.GetUnderlyingProperties().Application = InternalConst.ApplicationName;
            props.ExtendedProperties.GetUnderlyingProperties().AppVersion = AppVersion.ToString(3);
            return workbook;
        }
        else
        {
            var workbook = new NPOI.HSSF.UserModel.HSSFWorkbook();
            var dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = setting.Company;
            dsi.Category = setting.Category;
            workbook.DocumentSummaryInformation = dsi;
            var si = PropertySetFactory.CreateSummaryInformation();
            si.Title = setting.Title;
            si.Subject = setting.Subject;
            si.Author = setting.Author;
            si.CreateDateTime = DateTime.Now;
            si.Comments = setting.Description;
            si.ApplicationName = InternalConst.ApplicationName;
            workbook.SummaryInformation = si;
            return workbook;
        }
    }

    #endregion
        
}