using System;
using System.Collections;
using System.Linq;

namespace Bing.Offices.Parameters
{
    /// <summary>
    /// 导出参数基类
    /// </summary>
    public abstract class ExportParamBase
    {
        /// <summary>
        /// 对象类型
        /// </summary>
        public Type ObjectType { get; private set; }

        /// <summary>
        /// 记录数
        /// </summary>
        public int RecordCount { get; private set; }

        /// <summary>
        /// 数据源类型
        /// </summary>
        public DataSourceType DataSourceType { get; private set; }

        /// <summary>
        /// 数据源
        /// </summary>
        private object _dataSource;

        /// <summary>
        /// 数据源
        /// </summary>
        public object DataSource
        {
            get => _dataSource;
            set => InitDataSource(value);
        }

        /// <summary>
        /// 初始化数据源
        /// </summary>
        /// <param name="value">值</param>
        private void InitDataSource(object value)
        {
            _dataSource = value;
            if (_dataSource != null)
            {
                if (value is IList list)
                {
                    // 如果数据源是数组或列表
                    var tmpDataSource = list.Cast<object>().ToList();
                    if (tmpDataSource.Any())
                    {
                        ObjectType = tmpDataSource.First().GetType();
                        DataSourceType = DataSourceType.List;
                        RecordCount = tmpDataSource.Count;
                    }
                    else
                    {
                        _dataSource = null;
                    }
                }
                else
                {
                    // 如果数据源是一个基础对象
                    ObjectType = _dataSource.GetType();
                    DataSourceType = DataSourceType.BasicObject;
                    RecordCount = 1;
                }
            }
        }
    }
}
