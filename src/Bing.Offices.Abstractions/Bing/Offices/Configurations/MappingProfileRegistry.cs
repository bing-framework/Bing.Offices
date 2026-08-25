namespace Bing.Offices.Configurations;

/// <summary>
/// 线程安全的 Mapping Profile 注册表。
/// </summary>
public sealed class MappingProfileRegistry : IMappingProfileRegistry
{
    private readonly Dictionary<MappingProfileKey, ProfileDescriptor> _descriptors =
        new Dictionary<MappingProfileKey, ProfileDescriptor>();
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

    private static void ValidateKey(string profileName, Type modelType)
    {
        if (string.IsNullOrWhiteSpace(profileName))
            throw new ArgumentException("Profile 名称不能为空。", nameof(profileName));
        if (modelType == null)
            throw new ArgumentNullException(nameof(modelType));
    }

    private readonly struct MappingProfileKey : IEquatable<MappingProfileKey>
    {
        public MappingProfileKey(string name, MappingDirection direction, Type modelType)
        {
            Name = name;
            Direction = direction;
            ModelType = modelType;
        }

        private string Name { get; }
        private MappingDirection Direction { get; }
        private Type ModelType { get; }

        public bool Equals(MappingProfileKey other) => string.Equals(Name, other.Name,
            StringComparison.OrdinalIgnoreCase) && Direction == other.Direction && ModelType == other.ModelType;

        public override bool Equals(object obj) => obj is MappingProfileKey other && Equals(other);

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
