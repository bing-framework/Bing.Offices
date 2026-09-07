using Bing.Offices.Configurations;

namespace Bing.Offices.ProfileFixtures;

/// <summary>供外部程序集 Profile 扫描测试使用的双向映射 Profile。</summary>
public sealed class ExternalMappingProfile : IMappingProfile<ExternalImportModel, ExternalExportModel>
{
    /// <summary>配置外部输入和输出字段标题。</summary>
    public void Configure(FluentSetting<ExternalImportModel, ExternalExportModel> setting)
    {
        setting.Import.Property(model => model.Code).HasHeader("外部输入");
        setting.Export.Property(model => model.Label).HasHeader("外部输出");
    }
}

/// <summary>外部导入模型。</summary>
public sealed class ExternalImportModel
{
    /// <summary>外部输入编码。</summary>
    public string Code { get; set; }
}

/// <summary>外部导出模型。</summary>
public sealed class ExternalExportModel
{
    /// <summary>外部输出标签。</summary>
    public string Label { get; set; }
}
