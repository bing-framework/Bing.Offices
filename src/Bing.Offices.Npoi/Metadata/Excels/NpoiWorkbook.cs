using System;
using System.Collections.Generic;
using System.Linq;
using Bing.Offices.Abstractions.Metadata.Excels;
using Bing.Offices.Metadata.Excels;

namespace Bing.Offices.Npoi.Metadata.Excels
{
    /// <summary>
    /// Npoi工作簿
    /// </summary>
    public class NpoiWorkbook : IWorkbook
    {
        #region 字段

        /// <summary>
        /// 工作表名称最大长度
        /// </summary>
        private const int MaxSensitveSheetNameLength = 31;

        #endregion

        #region 属性

        /// <summary>
        /// 工作表列表
        /// </summary>
        public IList<IWorkSheet> Sheets { get; }

        /// <summary>
        /// 工作表
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        public IWorkSheet this[int sheetIndex] => Sheets[sheetIndex];

        #endregion

        #region 构造函数

        /// <summary>
        /// 初始化一个<see cref="NpoiWorkbook"/>类型的实例
        /// </summary>
        public NpoiWorkbook()
        {
            Sheets = new List<IWorkSheet>();
        }

        #endregion

        #region GetSheet(获取工作表)

        /// <summary>
        /// 获取工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        public IWorkSheet GetSheet(string sheetName) => Sheets.FirstOrDefault(sheet =>
            sheetName.Equals(sheet.Name, StringComparison.InvariantCultureIgnoreCase));

        #endregion

        #region GetSheetAt(获取工作表)

        /// <summary>
        /// 获取工作表
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        public IWorkSheet GetSheetAt(int sheetIndex)
        {
            ValidateSheetIndex(sheetIndex);
            return Sheets[sheetIndex];
        }

        /// <summary>
        /// 校验工作表索引
        /// </summary>
        /// <param name="index">索引</param>
        private void ValidateSheetIndex(int index)
        {
            var lastSheetIndex = Sheets.Count - 1;
            if (index < 0 || index > lastSheetIndex)
                throw new IndexOutOfRangeException($"工作表索引[{index}]已超出索引范围(0..{lastSheetIndex})");
        }

        #endregion

        #region CreateSheet(创建工作表)

        /// <summary>
        /// 创建工作表
        /// </summary>
        public IWorkSheet CreateSheet()
        {
            var sheetName = $"Sheet{Sheets.Count}";
            var index = 0;
            while (GetSheet(sheetName) != null)
            {
                sheetName = $"Sheet{index}";
                index++;
            }
            return CreateSheet(sheetName, 0);
        }

        /// <summary>
        /// 创建工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        public IWorkSheet CreateSheet(string sheetName) => CreateSheet(sheetName, 0);

        /// <summary>
        /// 创建工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="startRowIndex">起始行索引</param>
        public IWorkSheet CreateSheet(string sheetName, int startRowIndex)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
                throw new ArgumentNullException(nameof(sheetName));
            if (ContainsSheet(sheetName, Sheets.Count))
                throw new ArgumentException($"工作簿已存在该[{sheetName}]名称的工作表");
            if (sheetName.Length > 31)
                sheetName = sheetName.Substring(0, MaxSensitveSheetNameLength);
            ValidateSheetName(sheetName);
            IWorkSheet sheet = new WorkSheet(startRowIndex);
            sheet.Name = sheetName;
            Sheets.Add(sheet);
            return sheet;
        }

        /// <summary>
        /// 校验工作表名称
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        private void ValidateSheetName(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
                throw new ArgumentNullException(nameof(sheetName));
            var len = sheetName.Length;
            if (len < 1 || len > 31)
                throw new IndexOutOfRangeException($"工作表名称'{sheetName}'无效，字符数范围[1,31]");
            for (var i = 0; i < len; i++)
            {
                char ch = sheetName[i];
                switch (ch)
                {
                    case '/':
                    case '\\':
                    case '?':
                    case '*':
                    case ']':
                    case '[':
                    case ':':
                        break;
                    default:
                        continue;
                }
                throw new ArgumentException($"'{sheetName}'存在无效字符({ch})");
            }
            if (sheetName[0] == '\'' || sheetName[len - 1] == '\'')
                throw new ArgumentException($"无效工作表名称'{sheetName}'。工作表名称不能存在成对(')。");
        }

        /// <summary>
        /// 是否包含工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="excludeSheetIndex">忽略索引</param>
        private bool ContainsSheet(string sheetName, int excludeSheetIndex)
        {
            if (sheetName.Length > MaxSensitveSheetNameLength)
                sheetName = sheetName.Substring(0, MaxSensitveSheetNameLength);
            for (var i = 0; i < Sheets.Count; i++)
            {
                var name = Sheets[i].Name;
                if (sheetName.Length > MaxSensitveSheetNameLength)
                    name = name.Substring(0, MaxSensitveSheetNameLength);
                if (excludeSheetIndex != i && sheetName.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }
            return false;
        }

        #endregion

        #region AddSheet(添加工作表)

        /// <summary>
        /// 添加工作表
        /// </summary>
        /// <param name="sheet">工作表</param>
        public void AddSheet(IWorkSheet sheet)
        {
            if (sheet == null)
                return;
            if (Sheets.Any(x => x.Name == sheet.Name))
                return;
            Sheets.Add(sheet);
        }

        #endregion
    }
}
