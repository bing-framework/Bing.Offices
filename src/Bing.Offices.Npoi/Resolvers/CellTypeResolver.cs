using Bing.Offices.Metadata;
using NModel = NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Resolvers
{
    /// <summary>
    /// 单元格类型解析器
    /// </summary>
    internal class CellTypeResolver
    {
        /// <summary>
        /// 解析
        /// </summary>
        /// <param name="cellType">单元格类型</param>
        public static NModel.CellType Resolve(CellType cellType)
        {
            switch (cellType)
            {
                case CellType.Unknown:
                    return NModel.CellType.Unknown;
                case CellType.Numeric:
                    return NModel.CellType.Numeric;
                case CellType.String:
                    return NModel.CellType.String;
                case CellType.Formula:
                    return NModel.CellType.Formula;
                case CellType.Blank:
                    return NModel.CellType.Blank;
                case CellType.Boolean:
                    return NModel.CellType.Boolean;
                case CellType.Error:
                    return NModel.CellType.Error;
                default:
                    return NModel.CellType.Unknown;
            }
        }

        /// <summary>
        /// 解析
        /// </summary>
        /// <param name="cellType">单元格类型</param>
        public static CellType Resolve(NModel.CellType cellType)
        {
            switch (cellType)
            {
                case NModel.CellType.Unknown:
                    return CellType.Unknown;
                case NModel.CellType.Numeric:
                    return CellType.Numeric;
                case NModel.CellType.String:
                    return CellType.String;
                case NModel.CellType.Formula:
                    return CellType.Formula;
                case NModel.CellType.Blank:
                    return CellType.Blank;
                case NModel.CellType.Boolean:
                    return CellType.Boolean;
                case NModel.CellType.Error:
                    return CellType.Error;
                default:
                    return CellType.Unknown;
            }
        }
    }
}
