using System;
using Bing.Offices.Metadata;
using Bing.Offices.Metadata.Excels;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Resolvers;
using NModel = NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Metadata.Excels
{
    /// <summary>
    /// Npoi单元格
    /// </summary>
    internal class NpoiCell : ICell
    {
        /// <summary>
        /// 单元格
        /// </summary>
        private readonly NModel.ICell _cell;

        /// <summary>
        /// 初始化一个<see cref="NpoiCell"/>类型的实例
        /// </summary>
        /// <param name="cell">单元格</param>
        public NpoiCell(NModel.ICell cell) => _cell = cell ?? throw new ArgumentNullException(nameof(cell));

        /// <summary>
        /// 单元格类型
        /// </summary>
        public CellType CellType
        {
            get => CellTypeResolver.Resolve(_cell.CellType);
            set => _cell.SetCellType(CellTypeResolver.Resolve(value));
        }

        /// <summary>
        /// 值
        /// </summary>
        public object Value
        {
            get => _cell.GetValue(); 
            set => _cell.SetCellValue(value);
        }

        /// <summary>
        /// 行
        /// </summary>
        public IRow Row { get; set; }

        /// <summary>
        /// 列跨度
        /// </summary>
        public int ColumnSpan { get; set; }

        /// <summary>
        /// 行跨度
        /// </summary>
        public int RowSpan { get; }

        /// <summary>
        /// 列索引
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// 行索引
        /// </summary>
        public int RowIndex { get; }

        /// <summary>
        /// 物理行索引
        /// </summary>
        public int PhysicalRowIndex { get; }

        /// <summary>
        /// 结束列索引
        /// </summary>
        public int EndColumnIndex { get; }

        /// <summary>
        /// 结束行索引
        /// </summary>
        public int EndRowIndex { get; }

        /// <summary>
        /// 是否需要合并单元格。true:是,false:否
        /// </summary>
        public bool NeedMerge { get; }

        /// <summary>
        /// 名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 属性名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// 是否动态单元格
        /// </summary>
        public bool IsDynamic { get; set; }

        /// <summary>
        /// 是否为空单元格
        /// </summary>
        public bool IsNull()
        {
            throw new NotImplementedException();
        }
    }
}
