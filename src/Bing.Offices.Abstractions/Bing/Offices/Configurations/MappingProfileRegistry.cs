namespace Bing.Offices.Configurations;

/// <summary>
/// 线程安全的 Mapping Profile 注册表。
/// </summary>
public sealed class MappingProfileRegistry : IMappingProfileRegistry
{
    private readonly Dictionary<MappingProfileKey, IMappingProfileSnapshot> _profiles =
        new Dictionary<MappingProfileKey, IMappingProfileSnapshot>();
    private readonly object _sync = new object();

    /// <summary>
    /// 注册一个双向 Profile 快照。
    /// </summary>
    public void Register(string profileName, IMappingProfileSnapshot snapshot)
    {
        if (string.IsNullOrWhiteSpace(profileName))
            throw new ArgumentException("Profile 名称不能为空。", nameof(profileName));
        if (snapshot == null)
            throw new ArgumentNullException(nameof(snapshot));
        var importKey = new MappingProfileKey(profileName, MappingDirection.Import, snapshot.ImportType);
        var exportKey = new MappingProfileKey(profileName, MappingDirection.Export, snapshot.ExportType);
        lock (_sync)
        {
            if (_profiles.ContainsKey(importKey) || _profiles.ContainsKey(exportKey))
                throw new InvalidOperationException($"Profile 注册键重复: {profileName}");
            _profiles.Add(importKey, snapshot);
            _profiles.Add(exportKey, snapshot);
        }
    }

    /// <inheritdoc />
    public ExcelMappingProfile<TImport, TExport> Get<TImport, TExport>(string profileName)
        where TImport : class, new()
        where TExport : class, new()
    {
        if (!TryGet(profileName, MappingDirection.Import, typeof(TImport), out var importSnapshot)
            || !TryGet(profileName, MappingDirection.Export, typeof(TExport), out var exportSnapshot)
            || !ReferenceEquals(importSnapshot, exportSnapshot))
            throw new KeyNotFoundException($"未找到 Profile: {profileName}");
        return new ExcelMappingProfile<TImport, TExport>(importSnapshot.ImportConfiguration,
            exportSnapshot.ExportConfiguration);
    }

    /// <inheritdoc />
    public bool TryGet(string profileName, MappingDirection direction, Type modelType,
        out IMappingProfileSnapshot snapshot)
    {
        if (string.IsNullOrWhiteSpace(profileName))
            throw new ArgumentException("Profile 名称不能为空。", nameof(profileName));
        if (modelType == null)
            throw new ArgumentNullException(nameof(modelType));
        lock (_sync)
            return _profiles.TryGetValue(new MappingProfileKey(profileName, direction, modelType), out snapshot);
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
