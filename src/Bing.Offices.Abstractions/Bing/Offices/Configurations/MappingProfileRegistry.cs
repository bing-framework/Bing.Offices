namespace Bing.Offices.Configurations;

using System.ComponentModel;

/// <summary>
/// 线程安全的 Mapping Profile 注册表。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class MappingProfileRegistry : IMappingProfileRegistry
{
    /// <summary>按模型类型、方向和名称保存已注册的 Profile 描述。</summary>
    private readonly Dictionary<MappingProfileKey, ProfileDescriptor> _descriptors =
        new Dictionary<MappingProfileKey, ProfileDescriptor>();
    /// <summary>保护 Profile 注册表读写一致性的锁对象。</summary>
    private readonly object _sync = new object();

    /// <inheritdoc />
    public void Register(ProfileDescriptor descriptor)
    {
        if (descriptor == null)
            throw new ArgumentNullException(nameof(descriptor));
        var key = new MappingProfileKey(descriptor.Name, descriptor.Direction, descriptor.ModelType);
        lock (_sync)
        {
            if (_descriptors.ContainsKey(key))
                throw new InvalidOperationException($"Profile 注册键重复: {descriptor.Name}, {descriptor.Direction}, {descriptor.ModelType.FullName}");
            _descriptors.Add(key, descriptor);
        }
    }

    /// <inheritdoc />
    public bool TryGetDescriptor(string profileName, MappingDirection direction, Type modelType,
        out ProfileDescriptor descriptor)
    {
        ValidateKey(profileName, modelType);
        lock (_sync)
            return _descriptors.TryGetValue(new MappingProfileKey(profileName, direction, modelType),
                out descriptor);
    }

    /// <summary>验证 Profile 查找键的名称和模型类型。</summary>
    /// <param name="profileName">Profile 名称。</param>
    /// <param name="modelType">Profile 对应的模型类型。</param>
    private static void ValidateKey(string profileName, Type modelType)
    {
        if (string.IsNullOrWhiteSpace(profileName))
            throw new ArgumentException("Profile 名称不能为空。", nameof(profileName));
        if (modelType == null)
            throw new ArgumentNullException(nameof(modelType));
    }

    /// <summary>标识 Profile 注册项的复合键。</summary>
    private readonly struct MappingProfileKey : IEquatable<MappingProfileKey>
    {
        /// <summary>初始化一个 <see cref="MappingProfileKey" /> 类型的实例。</summary>
        /// <param name="name">Profile 名称。</param>
        /// <param name="direction">映射方向。</param>
        /// <param name="modelType">Profile 对应的模型类型。</param>
        public MappingProfileKey(string name, MappingDirection direction, Type modelType)
        {
            Name = name;
            Direction = direction;
            ModelType = modelType;
        }

        /// <summary>获取Profile 名称。</summary>
        private string Name { get; }
        /// <summary>获取映射方向。</summary>
        private MappingDirection Direction { get; }
        /// <summary>获取Profile 对应的模型类型。</summary>
        private Type ModelType { get; }

        /// <inheritdoc />
        public bool Equals(MappingProfileKey other) => string.Equals(Name, other.Name,
            StringComparison.OrdinalIgnoreCase) && Direction == other.Direction && ModelType == other.ModelType;

        /// <inheritdoc />
        public override bool Equals(object obj) => obj is MappingProfileKey other && Equals(other);

        /// <inheritdoc />
        public override int GetHashCode()
        {
            unchecked
            {
                var hash = StringComparer.OrdinalIgnoreCase.GetHashCode(Name);
                hash = hash * 397 ^ (int)Direction;
                hash = hash * 397 ^ ModelType.GetHashCode();
                return hash;
            }
        }
    }
}
