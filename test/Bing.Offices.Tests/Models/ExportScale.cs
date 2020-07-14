using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models
{
    /// <summary>
    /// 导出小数位
    /// </summary>
    public class ExportScale
    {
        /// <summary>
        /// 系统标识
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// byte
        /// </summary>
        [DecimalScale(1)]
        public byte Byte { get; set; }

        /// <summary>
        /// byte?
        /// </summary>
        [DecimalScale(1)]
        public byte? NullableByte { get; set; }

        /// <summary>
        /// short
        /// </summary>
        [DecimalScale(2)]
        public short Short { get; set; }

        /// <summary>
        /// short
        /// </summary>
        [DecimalScale(2)]
        public short? NullableShort { get; set; }

        /// <summary>
        /// ushort
        /// </summary>
        [DecimalScale(2)]
        public ushort UShort { get; set; }

        /// <summary>
        /// ushort?
        /// </summary>
        [DecimalScale(2)]
        public ushort? NullableUShort { get; set; }

        /// <summary>
        /// int
        /// </summary>
        [DecimalScale(3)]
        public int Int { get; set; }

        /// <summary>
        /// int?
        /// </summary>
        [DecimalScale(3)]
        public int? NullableInt { get; set; }

        /// <summary>
        /// uint
        /// </summary>
        [DecimalScale(3)]
        public uint UInt { get; set; }

        /// <summary>
        /// uint?
        /// </summary>
        [DecimalScale(3)] 
        public uint? NullableUInt { get; set; }

        /// <summary>
        /// long
        /// </summary>
        [DecimalScale(4)] 
        public long Long { get; set; }

        /// <summary>
        /// long?
        /// </summary>
        [DecimalScale(4)] 
        public long? NullableLong { get; set; }

        /// <summary>
        /// ulong
        /// </summary>
        [DecimalScale(4)] 
        public ulong ULong { get; set; }

        /// <summary>
        /// ulong?
        /// </summary>
        [DecimalScale(4)] 
        public ulong? NullableULong { get; set; }

        /// <summary>
        /// float
        /// </summary>
        [DecimalScale(1)] 
        public float Float { get; set; }

        /// <summary>
        /// float?
        /// </summary>
        [DecimalScale(2)]
        public float? NullableFloat { get; set; }

        /// <summary>
        /// double
        /// </summary>
        [DecimalScale(3)]
        public double Double { get; set; }

        /// <summary>
        /// double?
        /// </summary>
        [DecimalScale(4)]
        public double? NullableDouble { get; set; }

        /// <summary>
        /// decimal
        /// </summary>
        [DecimalScale(5)]
        public decimal Decimal { get; set; }

        /// <summary>
        /// decimal?
        /// </summary>
        [DecimalScale(7)]
        public decimal? NullableDecimal { get; set; }
    }
}
