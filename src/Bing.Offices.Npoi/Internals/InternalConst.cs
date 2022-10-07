namespace Bing.Offices.Npoi.Internals
{
    /// <summary>
    /// 内部常量
    /// </summary>
    internal static class InternalConst
    {
        /// <summary>
        /// Xls 最大工作表数量
        /// </summary>
        public const int MaxSheetCountXls = 256;

        /// <summary>
        /// Xlsx 最大工作表数量
        /// </summary>
        public const int MaxSheetCountXlsx = 16_384;

        /// <summary>
        /// Xls 最大行数
        /// </summary>
        public const int MaxRowCountXls = 65_536;

        /// <summary>
        /// Xlsx 最大行数
        /// </summary>
        public const int MaxRowCountXlsx = 1_048_576;

        /// <summary>
        /// 应用程序名称
        /// </summary>
        public const string ApplicationName = "Bing.Offices.Npoi";
    }
}