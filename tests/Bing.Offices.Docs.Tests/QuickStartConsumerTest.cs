using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Bing.Offices.Csv;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Emit;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Docs.Tests;

/// <summary>
/// 验证快速开始中的完整消费者程序能独立编译并执行。
/// </summary>
public sealed class QuickStartConsumerTest
{
    /// <summary>
    /// 测试 - 快速开始的 CSV 和 Excel 程序应按原始代码作为 ConsoleApplication 执行。
    /// </summary>
    [Fact]
    public async Task QuickStartPrograms_FromMarkdown_ShouldCompileAndExecuteIndividually()
    {
        var programs = ExtractPrograms().ToArray();
        Assert.Equal(2, programs.Length);

        for (var index = 0; index < programs.Length; index++)
        {
            var program = programs[index];
            var syntaxTree = CSharpSyntaxTree.ParseText(program.Code,
                new CSharpParseOptions(LanguageVersion.Latest));
            var mainType = FindMainType(syntaxTree);
            var requiresNpoi = program.Code.Contains("AddBingOfficesNpoi", StringComparison.Ordinal);
            var references = GetCompilationReferences(requiresNpoi).ToArray();
            if (!requiresNpoi)
                Assert.DoesNotContain(references, reference => IsNpoiReference(reference.Display ?? string.Empty));
            var compilation = CSharpCompilation.Create($"QuickStartConsumer_{index + 1}",
                new[] { syntaxTree }, references,
                new CSharpCompilationOptions(OutputKind.ConsoleApplication, mainTypeName: mainType));
            using var assemblyStream = new MemoryStream();
            EmitResult emit = compilation.Emit(assemblyStream);
            var errors = string.Join(Environment.NewLine, emit.Diagnostics
                .Where(diagnostic => diagnostic.Severity == DiagnosticSeverity.Error));
            Assert.True(emit.Success,
                $"quick-start.md fence {index + 1} 编译失败:{Environment.NewLine}{errors}");

            assemblyStream.Position = 0;
            var assembly = Assembly.Load(assemblyStream.ToArray());
            var main = assembly.GetType(mainType, throwOnError: true)!
                .GetMethod("Main", BindingFlags.Public | BindingFlags.Static)!;
            var execution = main.Invoke(null, null);
            await Assert.IsAssignableFrom<Task>(execution);
        }
    }

    private static IEnumerable<QuickStartProgram> ExtractPrograms()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Docs", "quick-start.md");
        var text = File.ReadAllText(path, System.Text.Encoding.UTF8);
        var fence = new string((char)96, 3);
        var escapedFence = System.Text.RegularExpressions.Regex.Escape(fence);
        var pattern = $"{escapedFence}csharp\\s*\\r?\\n(?<code>.*?)\\r?\\n[ ]*{escapedFence}";
        var matches = System.Text.RegularExpressions.Regex.Matches(text, pattern,
            System.Text.RegularExpressions.RegexOptions.Singleline
                | System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        for (var index = 0; index < matches.Count; index++)
            yield return new QuickStartProgram(matches[index].Groups["code"].Value);
    }

    private static string FindMainType(SyntaxTree syntaxTree)
    {
        var root = syntaxTree.GetRoot();
        var namespaceDeclaration = Assert.Single(root.DescendantNodes()
            .OfType<BaseNamespaceDeclarationSyntax>());
        var mainClass = Assert.Single(namespaceDeclaration.DescendantNodes()
            .OfType<ClassDeclarationSyntax>()
            .Where(type => type.Members.OfType<MethodDeclarationSyntax>()
                .Any(method => method.Identifier.ValueText == "Main")));
        return $"{namespaceDeclaration.Name}.{mainClass.Identifier.ValueText}";
    }

    private static IEnumerable<MetadataReference> GetCompilationReferences(bool includeNpoi)
    {
        var paths = ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES") ?? string.Empty)
            .Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries)
            .Concat(AppDomain.CurrentDomain.GetAssemblies().Select(assembly => assembly.Location))
            .Concat(new[]
            {
                typeof(ICsvImporter).Assembly.Location,
                typeof(BingOfficesServiceCollectionExtensions).Assembly.Location,
                typeof(ExcelSingleSheetImportExtensions).Assembly.Location,
                typeof(IServiceCollection).Assembly.Location,
                typeof(ServiceCollection).Assembly.Location
            })
            .Concat(includeNpoi
                ? new[] { typeof(ExcelNpoiServiceCollectionExtensions).Assembly.Location }
                : Array.Empty<string>())
            .Where(path => includeNpoi || !IsNpoiReference(path))
            .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
            .Distinct(StringComparer.OrdinalIgnoreCase);
        return paths.Select(path => MetadataReference.CreateFromFile(path));
    }

    private static bool IsNpoiReference(string path)
    {
        var name = Path.GetFileNameWithoutExtension(path);
        return string.Equals(name, "Bing.Offices.Npoi", StringComparison.OrdinalIgnoreCase)
            || name.StartsWith("NPOI", StringComparison.OrdinalIgnoreCase);
    }

    private sealed record QuickStartProgram(string Code);
}
