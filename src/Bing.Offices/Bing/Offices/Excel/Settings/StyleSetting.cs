using Bing.Offices.Styles;

namespace Bing.Offices.Excel.Settings
{
    /// <summary>
    /// 样式设置
    /// </summary>
    internal class StyleSetting
    {
        /// <summary>
        /// 字体大小
        /// </summary>
        public float? FontSize { get; set; }

        /// <summary>
        /// 字体名称
        /// </summary>
        public string FontName { get; set; } = "微软雅黑";

        /// <summary>
        /// 字体颜色
        /// </summary>
        public Color FontColor { get; set; } = Color.Black;

        /// <summary>
        /// 背景色
        /// </summary>
        public Color? BackgroundColor { get; set; }

        /// <summary>
        /// 宽度
        /// </summary>
        public float? Width { get; set; }

        /// <summary>
        /// 是否加粗
        /// </summary>
        public bool IsBold { get; set; }

        /// <summary>
        /// 是否允许自适应
        /// </summary>
        public bool AllowAutoFit { get; set; }

        /// <summary>
        /// 是否允许自动换行
        /// </summary>
        public bool AllowAutoWrap { get; set; }
    }
}
