namespace Bing.Offices.Parameters
{
    /// <summary>
    /// 导入参数
    /// </summary>
    public interface IImportParam : ITemplateParam
    {
        /// <summary>
        /// 数据填充方向
        /// </summary>
        DataDirection DataDirection { get; set; }
    }
}
