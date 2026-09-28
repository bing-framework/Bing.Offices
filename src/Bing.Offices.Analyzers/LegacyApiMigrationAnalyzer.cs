using System.Collections.Immutable;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;

namespace Bing.Offices.Analyzers;

/// <summary>
/// 提示已绑定的旧版 Excel 导出与动态列 API 的迁移入口。
/// </summary>
[DiagnosticAnalyzer(LanguageNames.CSharp)]
public sealed class LegacyApiMigrationAnalyzer : DiagnosticAnalyzer
{
    /// <summary>
    /// 旧版导出服务与选项诊断 ID。
    /// </summary>
    public const string LegacyExportId = "BOM001";

    /// <summary>
    /// 旧版动态列标记诊断 ID。
    /// </summary>
    public const string LegacyDynamicColumnId = "BOM002";

    /// <summary>
    /// 旧版导出服务与选项诊断定义。
    /// </summary>
    private static readonly DiagnosticDescriptor LegacyExport = new(
        LegacyExportId, "迁移旧版 Excel 导出入口",
        "'{0}' 是旧版导出 API；请改用 ExcelExport.Workbook 请求和 IExcelExporter，并人工迁移选项",
        "Bing.Offices.Migration", DiagnosticSeverity.Warning, true);

    /// <summary>
    /// 旧版动态列标记诊断定义。
    /// </summary>
    private static readonly DiagnosticDescriptor LegacyDynamicColumn = new(
        LegacyDynamicColumnId, "迁移旧版动态列标记",
        "'HasDynamicColumn' 是旧版动态列标记；多个字典请在 Entity 列表区域显式配置 DynamicColumnGroup",
        "Bing.Offices.Migration", DiagnosticSeverity.Warning, true);

    /// <inheritdoc />
    public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics =>
        ImmutableArray.Create(LegacyExport, LegacyDynamicColumn);

    /// <inheritdoc />
    public override void Initialize(AnalysisContext context)
    {
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
        context.EnableConcurrentExecution();
        context.RegisterSyntaxNodeAction(AnalyzeTypeReference, SyntaxKind.IdentifierName, SyntaxKind.GenericName);
        context.RegisterSyntaxNodeAction(AnalyzeAttribute, SyntaxKind.Attribute);
    }

    /// <summary>
    /// 识别已绑定到旧版导出服务或选项的类型引用。
    /// </summary>
    private static void AnalyzeTypeReference(SyntaxNodeAnalysisContext context)
    {
        var name = (SimpleNameSyntax)context.Node;
        if (name.Identifier.ValueText != "IExcelExportService"
            && name.Identifier.ValueText != "ExportOptions")
            return;
        var type = context.SemanticModel.GetSymbolInfo(name, context.CancellationToken).Symbol as INamedTypeSymbol;
        if (type == null || type.ContainingNamespace.ToDisplayString() != "Bing.Offices.Exports")
            return;

        var isLegacyExport = type.Name == "IExcelExportService" && type.Arity == 0
            || type.Name == "ExportOptions" && type.Arity == 1;
        if (isLegacyExport)
            context.ReportDiagnostic(Diagnostic.Create(LegacyExport, name.GetLocation(), type.Name));
    }

    /// <summary>
    /// 识别已绑定到旧版动态列标记的特性。
    /// </summary>
    private static void AnalyzeAttribute(SyntaxNodeAnalysisContext context)
    {
        var attribute = (AttributeSyntax)context.Node;
        var constructor = context.SemanticModel.GetSymbolInfo(attribute, context.CancellationToken).Symbol
            as IMethodSymbol;
        var type = constructor?.ContainingType;
        if (type?.Name == "HasDynamicColumnAttribute"
            && type.ContainingNamespace.ToDisplayString() == "Bing.Offices.Attributes")
            context.ReportDiagnostic(Diagnostic.Create(LegacyDynamicColumn, attribute.Name.GetLocation()));
    }
}
