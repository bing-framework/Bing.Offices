using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models;

/// <summary>
/// 导出小数位
/// </summary>
public class ExportScale
{
    /// <summary>
    /// 获取或设置系统标识。
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// 获取或设置 byte 值。
    /// </summary>
    [DecimalScale(1)]
    public byte Byte { get; set; }

    /// <summary>
    /// 获取或设置可空 byte 值。
    /// </summary>
    [DecimalScale(1)]
    public byte? NullableByte { get; set; }

    /// <summary>
    /// 获取或设置 short 值。
    /// </summary>
    [DecimalScale(2)]
    public short Short { get; set; }

    /// <summary>
    /// 获取或设置可空 short 值。
    /// </summary>
    [DecimalScale(2)]
    public short? NullableShort { get; set; }

    /// <summary>
    /// 获取或设置 ushort 值。
    /// </summary>
    [DecimalScale(2)]
    public ushort UShort { get; set; }

    /// <summary>
    /// 获取或设置可空 ushort 值。
    /// </summary>
    [DecimalScale(2)]
    public ushort? NullableUShort { get; set; }

    /// <summary>
    /// 获取或设置 int 值。
    /// </summary>
    [DecimalScale(3)]
    public int Int { get; set; }

    /// <summary>
    /// 获取或设置可空 int 值。
    /// </summary>
    [DecimalScale(3)]
    public int? NullableInt { get; set; }

    /// <summary>
    /// 获取或设置 uint 值。
    /// </summary>
    [DecimalScale(3)]
    public uint UInt { get; set; }

    /// <summary>
    /// 获取或设置可空 uint 值。
    /// </summary>
    [DecimalScale(3)]
    public uint? NullableUInt { get; set; }

    /// <summary>
    /// 获取或设置 long 值。
    /// </summary>
    [DecimalScale(4)]
    public long Long { get; set; }

    /// <summary>
    /// 获取或设置可空 long 值。
    /// </summary>
    [DecimalScale(4)]
    public long? NullableLong { get; set; }

    /// <summary>
    /// 获取或设置 ulong 值。
    /// </summary>
    [DecimalScale(4)]
    public ulong ULong { get; set; }

    /// <summary>
    /// 获取或设置可空 ulong 值。
    /// </summary>
    [DecimalScale(4)]
    public ulong? NullableULong { get; set; }

    /// <summary>
    /// 获取或设置 float 值。
    /// </summary>
    [DecimalScale(1)]
    public float Float { get; set; }

    /// <summary>
    /// 获取或设置可空 float 值。
    /// </summary>
    [DecimalScale(2)]
    public float? NullableFloat { get; set; }

    /// <summary>
    /// 获取或设置 double 值。
    /// </summary>
    [DecimalScale(3)]
    public double Double { get; set; }

    /// <summary>
    /// 获取或设置可空 double 值。
    /// </summary>
    [DecimalScale(4)]
    public double? NullableDouble { get; set; }

    /// <summary>
    /// 获取或设置 decimal 值。
    /// </summary>
    [DecimalScale(5)]
    public decimal Decimal { get; set; }

    /// <summary>
    /// 获取或设置可空 decimal 值。
    /// </summary>
    [DecimalScale(7)]
    public decimal? NullableDecimal { get; set; }
}
