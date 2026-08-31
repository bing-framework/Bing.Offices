using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using Bing.Offices.Configurations;

namespace Bing.Offices.Extensions;

internal static class ProfileDescriptorFactory
{
    /// <summary>单模型导入 Profile 的开放泛型契约类型。</summary>
    private static readonly Type ImportContract = typeof(IImportMappingProfile<>);
    /// <summary>单模型导出 Profile 的开放泛型契约类型。</summary>
    private static readonly Type ExportContract = typeof(IExportMappingProfile<>);
    /// <summary>同一模型双向 Profile 的开放泛型契约类型。</summary>
    private static readonly Type SameModelContract = typeof(IMappingProfile<>);
    /// <summary>不同导入和导出模型双向 Profile 的开放泛型契约类型。</summary>
    private static readonly Type DualModelContract = typeof(IMappingProfile<,>);

    /// <summary>确定类型是否声明至少一个受支持的映射 Profile 契约。</summary>
    /// <param name="profileType">待检查的 Profile 类型。</param>
    /// <returns>实现受支持契约时为 true。</returns>
    internal static bool HasSupportedContract(Type profileType) =>
        profileType.GetInterfaces().Any(IsSupportedContract);

    /// <summary>根据 Profile 类型的泛型契约生成方向和模型唯一键。</summary>
    /// <param name="profileType">声明映射契约的 Profile 类型。</param>
    /// <param name="name">Profile 注册名称。</param>
    /// <returns>按 Profile 契约生成的唯一键集合。</returns>
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

    /// <summary>执行 Profile 配置方法并创建所有方向的不可变描述符。</summary>
    /// <param name="profile">Profile 实例。</param>
    /// <param name="profileType">Profile 的运行时类型。</param>
    /// <param name="name">Profile 注册名称。</param>
    /// <returns>按实现契约生成的导入和导出描述符。</returns>
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

    /// <summary>执行单模型导入 Profile 并创建导入描述符。</summary>
    /// <param name="profile">Profile 实例。</param>
    /// <param name="profileType">Profile 的运行时类型。</param>
    /// <param name="name">Profile 注册名称。</param>
    /// <param name="modelType">导入模型类型。</param>
    /// <returns>导入方向的 Profile 描述符。</returns>
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

    /// <summary>执行单模型导出 Profile 并创建导出描述符。</summary>
    /// <param name="profile">Profile 实例。</param>
    /// <param name="profileType">Profile 的运行时类型。</param>
    /// <param name="name">Profile 注册名称。</param>
    /// <param name="modelType">导出模型类型。</param>
    /// <returns>导出方向的 Profile 描述符。</returns>
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

    /// <summary>执行同模型双向 Profile 并创建导入和导出描述符。</summary>
    /// <param name="profile">Profile 实例。</param>
    /// <param name="profileType">Profile 的运行时类型。</param>
    /// <param name="name">Profile 注册名称。</param>
    /// <param name="modelType">导入和导出共用的模型类型。</param>
    /// <returns>导入和导出方向的 Profile 描述符。</returns>
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

    /// <summary>执行双模型 Profile 并创建独立导入和导出描述符。</summary>
    /// <param name="profile">Profile 实例。</param>
    /// <param name="profileType">Profile 的运行时类型。</param>
    /// <param name="name">Profile 注册名称。</param>
    /// <param name="importType">导入模型类型。</param>
    /// <param name="exportType">导出模型类型。</param>
    /// <returns>导入和导出方向的 Profile 描述符。</returns>
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

    /// <summary>定位唯一兼容的接口 Configure 方法并保留其原始内部异常。</summary>
    /// <param name="profile">Profile 实例。</param>
    /// <param name="profileType">Profile 的运行时类型。</param>
    /// <param name="settingType">传入 Configure 方法的设置类型。</param>
    /// <param name="setting">传入 Configure 方法的设置实例。</param>
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

    /// <summary>确定接口 Configure 方法是否接受当前设置类型。</summary>
    /// <param name="contract">待检查的 Profile 契约接口。</param>
    /// <param name="settingType">当前设置类型。</param>
    /// <returns>Configure 方法接受设置类型时为 true。</returns>
    private static bool IsCompatibleSetting(Type contract, Type settingType)
    {
        var method = contract.GetMethod("Configure");
        return method != null && method.GetParameters()[0].ParameterType == settingType;
    }

    /// <summary>确定泛型接口是否为支持的导入、导出或双向 Profile 契约。</summary>
    /// <param name="contract">待检查的接口类型。</param>
    /// <returns>接口受当前工厂支持时为 true。</returns>
    private static bool IsSupportedContract(Type contract) => contract.IsGenericType
        && (contract.GetGenericTypeDefinition() == ImportContract
            || contract.GetGenericTypeDefinition() == ExportContract
            || contract.GetGenericTypeDefinition() == SameModelContract
            || contract.GetGenericTypeDefinition() == DualModelContract);

    /// <summary>创建在 Profile 名称、方向和模型类型范围内唯一的描述符键。</summary>
    /// <param name="name">Profile 注册名称。</param>
    /// <param name="direction">映射方向。</param>
    /// <param name="modelType">该方向使用的模型类型。</param>
    /// <returns>用于重复检测的稳定复合键。</returns>
    private static string CreateKey(string name, MappingDirection direction, Type modelType) =>
        string.Join("|", name, direction, modelType.FullName);
}
