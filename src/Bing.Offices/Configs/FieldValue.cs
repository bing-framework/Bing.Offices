namespace Bing.Offices.Configs
{
    /// <summary>
    /// Excel字段
    /// </summary>
    public class FieldValue
    {
        #region 导入导出通用

        /// <summary>
        /// 属性名。必填项
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// 标题。必填项
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 日期模式。如果设置的类型不是date，注册时，会抛出异常
        /// </summary>
        public string Pattern { get; set; }

        /// <summary>
        /// 格式化。例如(1:男,2:女)
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 单元格值转换器名称。解析Excel值接口定义：自定义实现（全类名）
        /// </summary>
        public string CellValueConverterName { get; set; }

        #endregion

        #region 导入

        /// <summary>
        /// 是否可以为null。仅导入时有效
        /// </summary>
        public bool IsNull { get; set; } = true;

        /// <summary>
        /// 正则表达式。仅导入时有效
        /// </summary>
        public string Regex { get; set; }

        /// <summary>
        /// 正则表达式不通过时，错误提示信息。仅导入时有效
        /// </summary>
        public string RegexErrMsg { get; set; }

        #endregion

        #region 导出

        /// <summary>
        /// 单元格宽度。仅导出时有效
        /// </summary>
        public int ColumnWidth { get; set; }

        /// <summary>
        /// 单元格对齐方式。支持：center,left,right。仅导出时有效
        /// </summary>
        public short Align { get; set; }

        /// <summary>
        /// 标题单元格背景色。仅导出时有效
        /// </summary>
        public short TitleBgColor { get; set; }

        /// <summary>
        /// 标题单元格字体色。仅导出时有效
        /// </summary>
        public short TitleFontColor { get; set; }

        /// <summary>
        /// 单元格样式是否与标题样式一致。仅导出时有效
        /// </summary>
        public bool UniformStyle { get; set; } = false;

        /// <summary>
        /// Decimal格式化规则。只对Number类型有效
        /// </summary>
        public string DecimalFormatPattern { get; set; }

        /// <summary>
        /// 默认值。当值为空时，字段的默认值。仅导出时有效
        /// </summary>
        public string DefaultValue { get; set; }

        #endregion

        /// <summary>
        /// 其他配置项。
        /// 与Excel导入导出无关,或许在自定义转换器时利用该参数可以更灵活配置一些其他信息,
        /// 比如一个转换器映射多个自动,配置该参数会更加灵活,可以配置成JSON等类型数据,具体根据自己的需求
        /// </summary>
        public string OtherConfig { get; set; }

        /// <summary>
        /// 初始化一个<see cref="FieldValue"/>类型的实例
        /// </summary>
        public FieldValue() { }

        /// <summary>
        /// 初始化一个<see cref="FieldValue"/>类型的实例
        /// </summary>
        /// <param name="propertyName">属性名</param>
        /// <param name="title">标题</param>
        /// <param name="pattern">日期模式</param>
        /// <param name="format">格式化</param>
        public FieldValue(string propertyName, string title, string pattern, string format)
        {
            PropertyName = propertyName;
            Title = title;
            Pattern = pattern;
            Format = format;
        }
    }
}
