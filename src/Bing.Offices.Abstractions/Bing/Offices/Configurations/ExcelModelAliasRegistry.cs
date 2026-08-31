using System;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace Bing.Offices.Configurations;

/// <summary>
/// 配置文档使用的业务模型别名注册表。
/// </summary>
public sealed class ExcelModelAliasRegistry
{
    /// <summary>按不区分大小写的业务别名保存已批准的模型和 Profile 身份。</summary>
    private readonly ConcurrentDictionary<string, ExcelModelAliasRegistration> _aliases =
        new ConcurrentDictionary<string, ExcelModelAliasRegistration>(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// 注册一个业务模型别名。
    /// </summary>
    public ExcelModelAliasRegistry Register(string alias)
    {
        ValidateAlias(alias);
        _aliases.TryAdd(alias, new ExcelModelAliasRegistration(null, null));
        return this;
    }

    /// <summary>
    /// 注册业务别名及其批准的模型/Profile 身份。
    /// </summary>
    public ExcelModelAliasRegistry Register(string alias, Type modelType, string profileName = null)
    {
        ValidateAlias(alias);
        if (modelType == null)
            throw new ArgumentNullException(nameof(modelType));
        if (!string.IsNullOrWhiteSpace(profileName))
            ValidateAlias(profileName);
        _aliases[alias] = new ExcelModelAliasRegistration(modelType, profileName);
        return this;
    }

    /// <summary>
    /// 判断别名是否已注册。
    /// </summary>
    public bool Contains(string alias) => !string.IsNullOrWhiteSpace(alias) && _aliases.ContainsKey(alias);

    /// <summary>
    /// 获取是否已配置 allowlist。
    /// </summary>
    public bool HasRegistrations => !_aliases.IsEmpty;

    /// <summary>
    /// 尝试解析业务别名的批准模型/Profile 身份。
    /// </summary>
    public bool TryResolve(string alias, out Type modelType, out string profileName)
    {
        modelType = null;
        profileName = null;
        if (string.IsNullOrWhiteSpace(alias) || !_aliases.TryGetValue(alias, out var registration))
            return false;
        modelType = registration.ModelType;
        profileName = registration.ProfileName;
        return true;
    }

    /// <summary>保存单个业务别名关联的可选模型类型和 Profile 名称。</summary>
    private sealed class ExcelModelAliasRegistration
    {
        /// <summary>创建模型别名注册记录。</summary>
        /// <param name="modelType">批准的模型类型；仅注册名称时可为空。</param>
        /// <param name="profileName">批准的 Profile 名称；未限制时可为空。</param>
        internal ExcelModelAliasRegistration(Type modelType, string profileName)
        {
            ModelType = modelType;
            ProfileName = profileName;
        }

        /// <summary>获取批准的模型类型。</summary>
        internal Type ModelType { get; }
        /// <summary>获取批准的 Profile 名称。</summary>
        internal string ProfileName { get; }
    }

    private static void ValidateAlias(string alias)
    {
        if (string.IsNullOrWhiteSpace(alias))
            throw new ArgumentException("模型别名不能为空。", nameof(alias));
        if (alias.Length > 256 || alias.IndexOfAny(new[] { '.', ',', '+', '[', ']' }) >= 0
            || alias.IndexOf("::", StringComparison.Ordinal) >= 0)
            throw new ArgumentException("模型别名必须是稳定业务别名，不能使用 CLR 或程序集限定类型名。", nameof(alias));
    }
}
