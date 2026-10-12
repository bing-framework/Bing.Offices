using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 为检查进程级 Forward 暂存文件的测试提供独占集合。
/// </summary>
[CollectionDefinition("ClosedXml Forward temporary files", DisableParallelization = true)]
public sealed class ClosedXmlForwardTemporaryFileCollection
{
}
