using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Xunit;

namespace Bing.Offices.Docs.Tests;

/// <summary>
/// 验证契约源码的能力目录和公开命名空间。
/// </summary>
public sealed class SourceLayoutTests
{
    /// <summary>
    /// 验证源码文件归入已定义的能力目录。
    /// </summary>
    [Theory]
    [InlineData("Entities")]
    [InlineData("Exports")]
    [InlineData("Imports")]
    public void AbstractionsSources_ShouldUseCapabilityDirectories(string area)
    {
        var root = GetAreaDirectory(area);
        var files = Directory.GetFiles(root, "*.cs", SearchOption.AllDirectories);
        Assert.NotEmpty(files);
        var folders = new Dictionary<string, string[]>(StringComparer.Ordinal)
        {
            ["Entities"] = new[] { "Layout", "Lists", "Import", "Export" },
            ["Exports"] = new[] { "Workbook", "Sheets", "DynamicColumns", "Reports", "Charts", "Content", "Streaming" },
            ["Imports"] = new[] { "Workbook", "Sheets", "Batch", "Failures", "Validation", "Resources" }
        };
        var rootFile = area switch
        {
            "Exports" => "IExcelExporter.cs",
            "Imports" => "IExcelImporter.cs",
            _ => null
        };

        foreach (var file in files)
        {
            var relative = Path.GetRelativePath(root, file);
            var parts = relative.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            Assert.True((parts.Length > 1 && folders[area].Contains(parts[0])) ||
                (parts.Length == 1 && parts[0] == rootFile), $"未归类的契约源码：{relative}");
        }
    }

    /// <summary>
    /// 验证公开契约保持原有命名空间。
    /// </summary>
    [Theory]
    [InlineData("Entities")]
    [InlineData("Exports")]
    [InlineData("Imports")]
    public void AbstractionsPublicTypes_ShouldKeepCompatibleNamespaces(string area)
    {
        var root = GetAreaDirectory(area);
        foreach (var file in Directory.EnumerateFiles(root, "*.cs", SearchOption.AllDirectories))
        {
            var source = File.ReadAllText(file, Encoding.UTF8);
            var syntax = CSharpSyntaxTree.ParseText(source).GetRoot();
            var publicTypes = syntax.DescendantNodes().OfType<MemberDeclarationSyntax>()
                .Where(member => member is BaseTypeDeclarationSyntax type && type.Modifiers.Any(token => token.IsKind(SyntaxKind.PublicKeyword)) ||
                    member is DelegateDeclarationSyntax declaration && declaration.Modifiers.Any(token => token.IsKind(SyntaxKind.PublicKeyword)));

            foreach (var type in publicTypes)
            {
                var declaredNamespace = type.Ancestors().OfType<BaseNamespaceDeclarationSyntax>().FirstOrDefault()?.Name.ToString();
                Assert.Equal($"Bing.Offices.{area}", declaredNamespace);
            }
        }
    }

    /// <summary>
    /// 验证 NPOI 实体执行类型使用统一的内部命名空间。
    /// </summary>
    [Theory]
    [InlineData("NpoiEntityLayoutSupport")]
    [InlineData("NpoiEntityExportExecutor")]
    [InlineData("NpoiEntityImportExecutor")]
    public void NpoiEntityExecutors_ShouldUseEntityNamespace(string typeName)
    {
        var file = Path.Combine(GetRepositoryDirectory(), "src", "Bing.Offices.Npoi", "Bing", "Offices",
            "Entities", typeName + ".cs");
        var source = File.ReadAllText(file, Encoding.UTF8);
        var syntax = CSharpSyntaxTree.ParseText(source).GetRoot();
        var type = Assert.Single(syntax.DescendantNodes().OfType<BaseTypeDeclarationSyntax>()
            .Where(declaration => declaration.Identifier.ValueText == typeName));

        Assert.Equal("Bing.Offices.Npoi.Entities",
            type.Ancestors().OfType<BaseNamespaceDeclarationSyntax>().FirstOrDefault()?.Name.ToString());
    }

    /// <summary>
    /// 获取契约能力目录。
    /// </summary>
    private static string GetAreaDirectory(string area)
    {
        var path = Path.Combine(GetRepositoryDirectory(), "src", "Bing.Offices.Abstractions", "Bing", "Offices", area);
        Assert.True(Directory.Exists(path), $"契约目录不存在：{path}");
        return path;
    }

    /// <summary>
    /// 获取仓库根目录。
    /// </summary>
    private static string GetRepositoryDirectory()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "Bing.Offices.sln")))
        {
            directory = directory.Parent;
        }

        Assert.NotNull(directory);
        return directory.FullName;
    }
}
