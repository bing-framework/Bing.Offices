using BenchmarkDotNet.Running;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 性能基准程序入口。
/// </summary>
public static class Program
{
    /// <summary>
    /// 运行流式 Excel 管线基准。
    /// </summary>
    /// <param name="args">BenchmarkDotNet 命令行参数。</param>
    public static void Main(string[] args) => BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
}
