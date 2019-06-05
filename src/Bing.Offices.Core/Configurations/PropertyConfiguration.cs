using Bing.Offices.Abstractions.Configurations;
using Bing.Offices.Abstractions.Settings;
using Bing.Offices.Settings;

namespace Bing.Offices.Configurations
{
    /// <summary>
    /// 属性配置
    /// </summary>
    internal class PropertyConfiguration : IPropertyConfiguration
    {
        /// <summary>
        /// 属性设置
        /// </summary>
        private readonly PropertySetting _propertySetting;

        /// <summary>
        /// 属性设置
        /// </summary>
        public IPropertySetting PropertySetting => _propertySetting;

        /// <summary>
        /// 初始化一个<see cref="PropertyConfiguration"/>类型的实例
        /// </summary>
        public PropertyConfiguration() : this(new PropertySetting())
        {
        }

        /// <summary>
        /// 初始化一个<see cref="PropertyConfiguration"/>类型的实例
        /// </summary>
        /// <param name="propertySetting">属性设置</param>
        public PropertyConfiguration(IPropertySetting propertySetting)
        {
            _propertySetting = propertySetting as PropertySetting;
        }

        /// <summary>
        /// 设置索引
        /// </summary>
        /// <param name="index">索引</param>
        public IPropertyConfiguration HasColumnIndex(int index)
        {
            if (index >= 0)
            {
                _propertySetting.Index = index;
            }

            return this;
        }

        /// <summary>
        /// 设置标题
        /// </summary>
        /// <param name="title">标题</param>
        public IPropertyConfiguration HasColumnTitle(string title)
        {
            if (!string.IsNullOrWhiteSpace(title))
            {
                _propertySetting.Title = title;
            }

            return this;
        }

        /// <summary>
        /// 设置格式化程序
        /// </summary>
        /// <param name="formatter">格式化程序</param>
        public IPropertyConfiguration HasColumnFormatter(string formatter)
        {
            if (!string.IsNullOrWhiteSpace(formatter))
            {
                _propertySetting.Formatter = formatter;
            }

            return this;
        }

        /// <summary>
        /// 设置忽略
        /// </summary>
        public IPropertyConfiguration Ignored()
        {
            _propertySetting.Ignored = true;
            return this;
        }

        /// <summary>
        /// 设置默认值
        /// </summary>
        /// <param name="value">默认值</param>
        public IPropertyConfiguration DefaultValue(object value)
        {
            if (value != null)
            {
                _propertySetting.DefaultValue = value;
            }

            return this;
        }
    }
}
