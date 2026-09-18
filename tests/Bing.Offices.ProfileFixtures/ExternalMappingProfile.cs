using Bing.Offices.Configurations;

namespace Bing.Offices.ProfileFixtures;

/// <summary>
/// 提供外部程序集 Profile 扫描测试使用的双向映射配置。
/// </summary>
public sealed class ExternalMappingProfile : IMappingProfile<ExternalImportModel, ExternalExportModel>
{
    /// <inheritdoc />
    /// <remarks>
    /// 将输入模型的 <see cref="ExternalImportModel.Code" /> 映射到“外部输入”标题，并将输出模型的 <see cref="ExternalExportModel.Label" /> 映射到“外部输出”标题。
    /// </remarks>
    public void Configure(FluentSetting<ExternalImportModel, ExternalExportModel> setting)
    {
        setting.Import.Property(model => model.Code).HasHeader("外部输入");
        setting.Export.Property(model => model.Label).HasHeader("外部输出");
    }
}

/// <summary>
/// 表示外部程序集扫描测试使用的导入模型。
/// </summary>
public sealed class ExternalImportModel
{
    /// <summary>
    /// 获取或设置外部输入编码。
    /// </summary>
    public string Code { get; set; }
}

/// <summary>
/// 表示外部程序集扫描测试使用的导出模型。
/// </summary>
public sealed class ExternalExportModel
{
    /// <summary>
    /// 获取或设置外部输出标签。
    /// </summary>
    public string Label { get; set; }
}
