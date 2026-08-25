using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using Bing.Offices.Configurations;

namespace Bing.Offices.Npoi.Extensions;

internal static class ProfileDescriptorFactory
{
    private static readonly Type ImportContract = typeof(IImportMappingProfile<>);
    private static readonly Type ExportContract = typeof(IExportMappingProfile<>);
    private static readonly Type SameModelContract = typeof(IMappingProfile<>);
    private static readonly Type DualModelContract = typeof(IMappingProfile<,>);

    internal static bool HasSupportedContract(Type profileType) =>
        profileType.GetInterfaces().Any(IsSupportedContract);

    internal static IReadOnlyList<string> GetKeys(Type profileType, string name)
    {
        if (profileType == null)
            throw new ArgumentNullException(nameof(profileType));
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Profile 名称不能为空。", nameof(name));
        var keys = new List<string>();
        foreach (var contract in profileType.GetInterfaces().Where(IsSupportedContract))
        {
            var definition = contract.GetGenericTypeDefinition();
            var arguments = contract.GetGenericArguments();
            if (definition == ImportContract)
                keys.Add(CreateKey(name, MappingDirection.Import, arguments[0]));
            else if (definition == ExportContract)
                keys.Add(CreateKey(name, MappingDirection.Export, arguments[0]));
            else if (definition == SameModelContract)
            {
                keys.Add(CreateKey(name, MappingDirection.Import, arguments[0]));
                keys.Add(CreateKey(name, MappingDirection.Export, arguments[0]));
            }
            else if (definition == DualModelContract)
            {
                keys.Add(CreateKey(name, MappingDirection.Import, arguments[0]));
                keys.Add(CreateKey(name, MappingDirection.Export, arguments[1]));
            }
        }
        if (keys.Distinct(StringComparer.OrdinalIgnoreCase).Count() != keys.Count)
            throw new InvalidOperationException(
                $"Profile 契约产生重复方向 descriptor: {profileType.FullName}。");
        return keys;
    }

    internal static IReadOnlyList<ProfileDescriptor> Create(object profile, Type profileType, string name)
    {
        if (profile == null)
            throw new ArgumentNullException(nameof(profile));
        var contracts = profileType.GetInterfaces()
            .Where(IsSupportedContract)
            .OrderBy(contract => contract.FullName, StringComparer.Ordinal)
            .ToArray();
        var descriptors = new List<ProfileDescriptor>();
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var contract in contracts)
        {
            var definition = contract.GetGenericTypeDefinition();
            var arguments = contract.GetGenericArguments();
            if (definition == ImportContract)
                descriptors.Add(CreateImport(profile, profileType, name, arguments[0]));
            else if (definition == ExportContract)
                descriptors.Add(CreateExport(profile, profileType, name, arguments[0]));
            else if (definition == SameModelContract)
                descriptors.AddRange(CreateSameModel(profile, profileType, name, arguments[0]));
            else if (definition == DualModelContract)
                descriptors.AddRange(CreateDualModel(profile, profileType, name, arguments[0], arguments[1]));
        }
        foreach (var descriptor in descriptors)
        {
            var key = string.Join("|", descriptor.Name, descriptor.Direction, descriptor.ModelType.FullName);
            if (!keys.Add(key))
                throw new InvalidOperationException(
                    $"Profile 契约产生重复方向 descriptor: {descriptor.Name}，方向: {descriptor.Direction}，模型: {descriptor.ModelType.FullName}。");
        }
        return descriptors;
    }

    private static ProfileDescriptor CreateImport(object profile, Type profileType, string name, Type modelType)
    {
        var builderType = typeof(ImportMappingBuilder<>).MakeGenericType(modelType);
        var builder = Activator.CreateInstance(builderType);
        InvokeConfigure(profile, profileType, builderType, builder);
        var configuration = (ExcelMappingConfiguration)builderType.GetMethod(nameof(ImportMappingBuilder<object>.Build))
            .Invoke(builder, new object[] { MappingSourceKind.Profile });
        configuration.Profile = name;
        return new ProfileDescriptor(name, MappingDirection.Import, modelType, configuration, profileType);
    }

    private static ProfileDescriptor CreateExport(object profile, Type profileType, string name, Type modelType)
    {
        var builderType = typeof(ExportMappingBuilder<>).MakeGenericType(modelType);
        var builder = Activator.CreateInstance(builderType);
        InvokeConfigure(profile, profileType, builderType, builder);
        var configuration = (ExcelMappingConfiguration)builderType.GetMethod(nameof(ExportMappingBuilder<object>.Build))
            .Invoke(builder, new object[] { MappingSourceKind.Profile });
        configuration.Profile = name;
        return new ProfileDescriptor(name, MappingDirection.Export, modelType, configuration, profileType);
    }

    private static IReadOnlyList<ProfileDescriptor> CreateSameModel(object profile, Type profileType,
        string name, Type modelType)
    {
        var settingType = typeof(FluentSetting<,>).MakeGenericType(modelType, modelType);
        var setting = Activator.CreateInstance(settingType);
        InvokeConfigure(profile, profileType, settingType, setting);
        var import = (ExcelMappingConfiguration)settingType.GetMethod(nameof(FluentSetting<object, object>.BuildImportConfiguration))
            .Invoke(setting, null);
        var export = (ExcelMappingConfiguration)settingType.GetMethod(nameof(FluentSetting<object, object>.BuildExportConfiguration))
            .Invoke(setting, null);
        import.Profile = name;
        export.Profile = name;
        return new[]
        {
            new ProfileDescriptor(name, MappingDirection.Import, modelType, import, profileType),
            new ProfileDescriptor(name, MappingDirection.Export, modelType, export, profileType)
        };
    }

    private static IReadOnlyList<ProfileDescriptor> CreateDualModel(object profile, Type profileType,
        string name, Type importType, Type exportType)
    {
        var settingType = typeof(FluentSetting<,>).MakeGenericType(importType, exportType);
        var setting = Activator.CreateInstance(settingType);
        InvokeConfigure(profile, profileType, settingType, setting);
        var import = (ExcelMappingConfiguration)settingType.GetMethod(nameof(FluentSetting<object, object>.BuildImportConfiguration))
            .Invoke(setting, null);
        var export = (ExcelMappingConfiguration)settingType.GetMethod(nameof(FluentSetting<object, object>.BuildExportConfiguration))
            .Invoke(setting, null);
        import.Profile = name;
        export.Profile = name;
        return new[]
        {
            new ProfileDescriptor(name, MappingDirection.Import, importType, import, profileType),
            new ProfileDescriptor(name, MappingDirection.Export, exportType, export, profileType)
        };
    }

    private static void InvokeConfigure(object profile, Type profileType, Type settingType, object setting)
    {
        var contracts = profileType.GetInterfaces().Where(item => item.IsGenericType
            && (item.GetGenericTypeDefinition() == ImportContract
                || item.GetGenericTypeDefinition() == ExportContract
                || item.GetGenericTypeDefinition() == SameModelContract
                || item.GetGenericTypeDefinition() == DualModelContract)
            && IsCompatibleSetting(item, settingType)).ToArray();
        if (contracts.Length != 1)
            throw new InvalidOperationException(
                $"Profile 配置方法不唯一: {profileType.FullName}，设置类型: {settingType.FullName}。");
        try
        {
            contracts[0].GetMethod("Configure").Invoke(profile, new[] { setting });
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }

    private static bool IsCompatibleSetting(Type contract, Type settingType)
    {
        var method = contract.GetMethod("Configure");
        return method != null && method.GetParameters()[0].ParameterType == settingType;
    }

    private static bool IsSupportedContract(Type contract) => contract.IsGenericType
        && (contract.GetGenericTypeDefinition() == ImportContract
            || contract.GetGenericTypeDefinition() == ExportContract
            || contract.GetGenericTypeDefinition() == SameModelContract
            || contract.GetGenericTypeDefinition() == DualModelContract);

    private static string CreateKey(string name, MappingDirection direction, Type modelType) =>
        string.Join("|", name, direction, modelType.FullName);
}
