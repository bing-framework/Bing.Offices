using Bing.Offices.Configurations;

namespace Bing.Offices.ProfileFixtures;

public sealed class ExternalMappingProfile : IMappingProfile<ExternalImportModel, ExternalExportModel>
{
    public void Configure(FluentSetting<ExternalImportModel, ExternalExportModel> setting)
    {
        setting.Import.Property(model => model.Code).HasHeader("外部输入");
        setting.Export.Property(model => model.Label).HasHeader("外部输出");
    }
}

public sealed class ExternalImportModel
{
    public string Code { get; set; }
}

public sealed class ExternalExportModel
{
    public string Label { get; set; }
}
