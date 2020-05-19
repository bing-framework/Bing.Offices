namespace Bing.Offices.Imports
{
    /// <summary>
    /// 验证结果
    /// </summary>
    public class ValidateResult
    {
        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 行号
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMsg { get; set; } = string.Empty;
    }
}
