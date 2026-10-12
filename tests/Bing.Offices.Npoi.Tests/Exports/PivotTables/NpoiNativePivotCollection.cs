using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 串行运行依赖原生 XSSF 透视初始化的职责测试。
/// </summary>
[CollectionDefinition("NpoiNativePivot", DisableParallelization = true)]
public sealed class NpoiNativePivotCollection
{
}
