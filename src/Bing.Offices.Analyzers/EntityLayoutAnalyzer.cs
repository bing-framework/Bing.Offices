using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Globalization;
using System.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Operations;

namespace Bing.Offices.Analyzers;

/// <summary>
/// 检查可静态确定的实体布局地址和动态列名称。
/// </summary>
[DiagnosticAnalyzer(LanguageNames.CSharp)]
public sealed class EntityLayoutAnalyzer : DiagnosticAnalyzer
{
    /// <summary>
    /// 固定单元格特性的完整类型名称。
    /// </summary>
    private const string CellAttributeName = "Bing.Offices.Entities.ExcelEntityCellAttribute";
    /// <summary>
    /// 实体布局构建器的完整类型名称。
    /// </summary>
    private const string LayoutBuilderName = "Bing.Offices.Entities.ExcelEntityLayoutBuilder<TEntity>";
    /// <summary>
    /// 列表区域构建器的完整类型名称。
    /// </summary>
    private const string RegionBuilderName = "Bing.Offices.Entities.ExcelEntityListRegionBuilder<TItem>";
    /// <summary>
    /// 动态列定义的完整类型名称。
    /// </summary>
    private const string DynamicDefinitionName = "Bing.Offices.Exports.ExcelDynamicColumnDefinition";
    /// <summary>
    /// 实体布局工厂的完整类型名称。
    /// </summary>
    private const string ExcelEntityName = "Bing.Offices.Entities.ExcelEntity";

    /// <summary>
    /// 无效的常量 A1 地址。
    /// </summary>
    public const string InvalidAddressId = "BOE001";
    /// <summary>
    /// 同一工作表上的重复固定地址。
    /// </summary>
    public const string DuplicateAddressId = "BOE002";
    /// <summary>
    /// 属性式单元格无法写回导入结果。
    /// </summary>
    public const string ReadOnlyPropertyId = "BOE003";
    /// <summary>
    /// 同一列表区域内重复的动态列名称。
    /// </summary>
    public const string DuplicateDynamicNameId = "BOE004";
    /// <summary>
    /// 同一工作表中可确定的布局区域交叉。
    /// </summary>
    public const string OverlappingAreaId = "BOE005";

    /// <summary>
    /// 无效地址诊断定义。
    /// </summary>
    private static readonly DiagnosticDescriptor InvalidAddress = new(
        InvalidAddressId, "实体布局地址无效", "地址 '{0}' 不是有效的 Excel A1 单元格地址",
        "Bing.Offices.EntityLayout", DiagnosticSeverity.Error, true);
    /// <summary>
    /// 重复固定地址诊断定义。
    /// </summary>
    private static readonly DiagnosticDescriptor DuplicateAddress = new(
        DuplicateAddressId, "实体布局地址重复", "工作表 '{0}' 的固定单元格地址 '{1}' 重复",
        "Bing.Offices.EntityLayout", DiagnosticSeverity.Error, true);
    /// <summary>
    /// 只读属性导入提示定义。
    /// </summary>
    private static readonly DiagnosticDescriptor ReadOnlyProperty = new(
        ReadOnlyPropertyId, "属性式单元格不能导入", "属性 '{0}' 没有 setter；可以导出，但导入无法写回",
        "Bing.Offices.EntityLayout", DiagnosticSeverity.Info, true);
    /// <summary>
    /// 重复动态列名称诊断定义。
    /// </summary>
    private static readonly DiagnosticDescriptor DuplicateDynamicName = new(
        DuplicateDynamicNameId, "动态列名称重复", "列表区域内的动态列 Key、标题、别名或组标识 '{0}' 重复",
        "Bing.Offices.EntityLayout", DiagnosticSeverity.Error, true);
    /// <summary>
    /// 布局区域交叉诊断定义。
    /// </summary>
    private static readonly DiagnosticDescriptor OverlappingArea = new(
        OverlappingAreaId, "实体布局区域重叠", "工作表 '{0}' 的布局区域 '{1}' 与已有区域重叠",
        "Bing.Offices.EntityLayout", DiagnosticSeverity.Error, true);

    /// <inheritdoc />
    public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics =>
        ImmutableArray.Create(InvalidAddress, DuplicateAddress, ReadOnlyProperty, DuplicateDynamicName,
            OverlappingArea);

    /// <inheritdoc />
    public override void Initialize(AnalysisContext context)
    {
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
        context.EnableConcurrentExecution();
        context.RegisterSymbolAction(AnalyzeProperty, SymbolKind.Property);
        context.RegisterSyntaxNodeAction(AnalyzeInvocation, SyntaxKind.InvocationExpression);
    }

    /// <summary>
    /// 检查属性式单元格的地址与可写性。
    /// </summary>
    private static void AnalyzeProperty(SymbolAnalysisContext context)
    {
        var property = (IPropertySymbol)context.Symbol;
        var attribute = GetCellAttribute(property);
        if (attribute == null || attribute.ConstructorArguments.Length < 2)
            return;
        var sheet = attribute.ConstructorArguments[0].Value as string;
        var address = attribute.ConstructorArguments[1].Value as string;
        var location = attribute.ApplicationSyntaxReference?.GetSyntax(context.CancellationToken).GetLocation()
            ?? property.Locations.FirstOrDefault();
        if (location == null)
            return;
        if (address == null || sheet == null)
            return;
        if (!TryNormalizeAddress(address, out var normalized))
        {
            context.ReportDiagnostic(Diagnostic.Create(InvalidAddress, location, address));
            return;
        }
        if (property.SetMethod == null)
            context.ReportDiagnostic(Diagnostic.Create(ReadOnlyProperty, location, property.Name));

        foreach (var other in ComparableProperties(property))
        {
            var otherAttribute = GetCellAttribute(other);
            if (otherAttribute == null || otherAttribute.ConstructorArguments.Length < 2)
                continue;
            var otherSheet = otherAttribute.ConstructorArguments[0].Value as string;
            var otherAddress = otherAttribute.ConstructorArguments[1].Value as string;
            if (string.Equals(sheet, otherSheet, StringComparison.OrdinalIgnoreCase)
                && otherAddress != null && TryNormalizeAddress(otherAddress, out var otherNormalized)
                && normalized == otherNormalized)
            {
                context.ReportDiagnostic(Diagnostic.Create(DuplicateAddress, location, sheet, address));
                break;
            }
        }
    }

    /// <summary>
    /// 获取属性上的实体固定单元格特性。
    /// </summary>
    private static AttributeData? GetCellAttribute(IPropertySymbol property) => property.GetAttributes()
        .FirstOrDefault(attribute => attribute.AttributeClass?.ToDisplayString() == CellAttributeName);

    /// <summary>
    /// 枚举当前及基类中可能占用同一地址的属性。
    /// </summary>
    private static IEnumerable<IPropertySymbol> ComparableProperties(IPropertySymbol property)
    {
        foreach (var other in property.ContainingType.GetMembers().OfType<IPropertySymbol>())
            if (string.CompareOrdinal(other.Name, property.Name) < 0)
                yield return other;
        for (var baseType = property.ContainingType.BaseType; baseType != null; baseType = baseType.BaseType)
            foreach (var other in baseType.GetMembers().OfType<IPropertySymbol>())
                if (other.DeclaredAccessibility == Accessibility.Public && other.Name != property.Name)
                    yield return other;
    }

    /// <summary>
    /// 分派已绑定到实体布局构建器的方法调用。
    /// </summary>
    private static void AnalyzeInvocation(SyntaxNodeAnalysisContext context)
    {
        var invocation = (InvocationExpressionSyntax)context.Node;
        var operation = context.SemanticModel.GetOperation(invocation, context.CancellationToken)
            as IInvocationOperation;
        if (operation == null)
            return;
        var owner = operation.TargetMethod.ContainingType.OriginalDefinition.ToDisplayString();
        if (owner == ExcelEntityName && (operation.TargetMethod.Name == "Layout"
            || operation.TargetMethod.Name == "LayoutFromAttributes"))
            AnalyzeLayoutAreas(context, operation);
        else if (operation.TargetMethod.Name == "Cell" && owner == LayoutBuilderName)
            AnalyzeFixedCell(context, invocation, operation);
        else if (operation.TargetMethod.Name == "DynamicColumnGroup" && owner == RegionBuilderName)
            AnalyzeDynamicGroup(context, invocation, operation);
    }

    /// <summary>
    /// 检查单条实体布局调用链中可确定的区域冲突。
    /// </summary>
    private static void AnalyzeLayoutAreas(SyntaxNodeAnalysisContext context, IInvocationOperation layout)
    {
        var configure = GetArgument(layout, "configure")?.Expression as LambdaExpressionSyntax;
        if (configure?.Body is not InvocationExpressionSyntax last)
            return;
        var calls = ReceiverChain(last).Reverse().Concat(new[] { last }).ToArray();
        var areas = new List<KnownArea>();
        var usesAttributes = layout.TargetMethod.Name == "LayoutFromAttributes" || calls.Any(call =>
        {
            var bound = context.SemanticModel.GetOperation(call, context.CancellationToken)
                as IInvocationOperation;
            return bound?.TargetMethod.Name == "CellsFromAttributes"
                && bound.TargetMethod.ContainingType.OriginalDefinition.ToDisplayString() == LayoutBuilderName;
        });
        if (usesAttributes && layout.TargetMethod.TypeArguments.FirstOrDefault() is INamedTypeSymbol entity)
            CollectAttributedCells(entity, areas);

        foreach (var call in calls)
        {
            if (context.SemanticModel.GetOperation(call, context.CancellationToken)
                is not IInvocationOperation bound
                || bound.TargetMethod.ContainingType.OriginalDefinition.ToDisplayString() != LayoutBuilderName)
                continue;
            var name = bound.TargetMethod.Name;
            if (name != "Cell" && name != "Merge" && name != "ListRegion")
                continue;
            var sheetArgument = GetArgument(bound, "sheetName");
            var addressArgument = GetArgument(bound, name == "Merge" ? "range"
                : name == "ListRegion" ? "startAddress" : "address");
            if (sheetArgument == null || addressArgument == null)
                continue;
            var sheet = ConstantString(context.SemanticModel, sheetArgument.Expression);
            var address = ConstantString(context.SemanticModel, addressArgument.Expression);
            if (sheet == null || address == null)
                continue;
            KnownArea? current = null;
            if (name == "Cell" && TryParseAddress(address, out var cellRow, out var cellColumn))
                current = new KnownArea(sheet, address, cellRow, cellRow, cellColumn, cellColumn, "Cell");
            else if (name == "Merge" && TryParseRange(address, out var merge))
                current = new KnownArea(sheet, address, merge.FirstRow, merge.LastRow,
                    merge.FirstColumn, merge.LastColumn, "Merge");
            else if (name == "ListRegion" && TryParseAddress(address, out var startRow,
                         out var startColumn) && TryGetRegionEnd(context, bound, out var endRow,
                         out var endColumn) && endRow >= startRow && endColumn >= startColumn)
                current = new KnownArea(sheet, address, startRow, endRow, startColumn, endColumn, "Region");
            if (current == null)
                continue;
            var conflicts = areas.Any(previous => Conflicts(previous, current, areas));
            if (name == "Merge" && areas.Where(previous => IsCell(previous)
                    && string.Equals(previous.Sheet, sheet, StringComparison.OrdinalIgnoreCase)
                    && current.Contains(previous.FirstRow, previous.FirstColumn))
                .Select(previous => (previous.FirstRow, previous.FirstColumn)).Distinct().Skip(1).Any())
                conflicts = true;
            if (conflicts)
                context.ReportDiagnostic(Diagnostic.Create(OverlappingArea,
                    addressArgument.GetLocation(), sheet, address));
            areas.Add(current);
        }
    }

    /// <summary>
    /// 收集属性式布局中公开实例属性的常量地址。
    /// </summary>
    private static void CollectAttributedCells(INamedTypeSymbol entity, ICollection<KnownArea> areas)
    {
        var names = new HashSet<string>(StringComparer.Ordinal);
        for (var type = entity; type != null; type = type.BaseType)
            foreach (var property in type.GetMembers().OfType<IPropertySymbol>())
            {
                if (property.DeclaredAccessibility != Accessibility.Public || property.IsStatic
                    || property.IsIndexer || !names.Add(property.Name))
                    continue;
                var attribute = GetCellAttribute(property);
                if (attribute?.ConstructorArguments.Length != 2
                    || attribute.ConstructorArguments[0].Value is not string sheet
                    || attribute.ConstructorArguments[1].Value is not string address
                    || !TryParseAddress(address, out var row, out var column))
                    continue;
                areas.Add(new KnownArea(sheet, address, row, row, column, column, "AttributeCell"));
            }
    }

    /// <summary>
    /// 获取内联列表区域的常量结束地址。
    /// </summary>
    private static bool TryGetRegionEnd(SyntaxNodeAnalysisContext context, IInvocationOperation region,
        out int row, out int column)
    {
        row = column = 0;
        var configure = GetArgument(region, "configure")?.Expression as LambdaExpressionSyntax;
        if (configure?.Body is not InvocationExpressionSyntax last)
            return false;
        var found = false;
        foreach (var call in ReceiverChain(last).Reverse().Concat(new[] { last }))
        {
            if (context.SemanticModel.GetOperation(call, context.CancellationToken)
                is not IInvocationOperation bound || bound.TargetMethod.Name != "End"
                || bound.TargetMethod.ContainingType.OriginalDefinition.ToDisplayString() != RegionBuilderName)
                continue;
            var address = GetArgument(bound, "address");
            if (address == null)
                return false;
            var value = ConstantString(context.SemanticModel, address.Expression);
            if (value == null || !TryParseAddress(value, out row, out column))
                return false;
            found = true;
        }
        return found;
    }

    /// <summary>
    /// 判断两项已知布局区域是否按构建器规则冲突。
    /// </summary>
    private static bool Conflicts(KnownArea left, KnownArea right, IReadOnlyList<KnownArea> areas)
    {
        if (!string.Equals(left.Sheet, right.Sheet, StringComparison.OrdinalIgnoreCase))
            return false;
        if (IsCell(left) && IsCell(right))
        {
            if (left.FirstRow == right.FirstRow && left.FirstColumn == right.FirstColumn)
                return left.Kind == "AttributeCell" || right.Kind == "AttributeCell";
            return areas.Any(merge => merge.Kind == "Merge"
                && string.Equals(merge.Sheet, left.Sheet, StringComparison.OrdinalIgnoreCase)
                && merge.Contains(left.FirstRow, left.FirstColumn)
                && merge.Contains(right.FirstRow, right.FirstColumn));
        }
        if ((IsCell(left) && right.Kind == "Merge")
            || (left.Kind == "Merge" && IsCell(right)))
            return false;
        return left.FirstRow <= right.LastRow && right.FirstRow <= left.LastRow
            && left.FirstColumn <= right.LastColumn && right.FirstColumn <= left.LastColumn;
    }

    /// <summary>
    /// 判断区域是否来自固定单元格声明。
    /// </summary>
    private static bool IsCell(KnownArea area) => area.Kind == "Cell"
        || area.Kind == "AttributeCell";

    /// <summary>
    /// 解析单个 XLSX A1 地址的行列号。
    /// </summary>
    private static bool TryParseAddress(string address, out int row, out int column)
    {
        row = column = 0;
        if (!TryNormalizeAddress(address, out var normalized))
            return false;
        var separator = normalized.IndexOf(':');
        return int.TryParse(normalized.Substring(0, separator), NumberStyles.None,
                   CultureInfo.InvariantCulture, out column)
            && int.TryParse(normalized.Substring(separator + 1), NumberStyles.None,
                CultureInfo.InvariantCulture, out row);
    }

    /// <summary>
    /// 解析单个常量 A1 矩形范围。
    /// </summary>
    private static bool TryParseRange(string range, out KnownArea area)
    {
        area = null!;
        var parts = range.Split(':');
        if (parts.Length > 2 || !TryParseAddress(parts[0], out var firstRow,
                out var firstColumn) || !TryParseAddress(parts[parts.Length - 1],
                out var lastRow, out var lastColumn)
            || lastRow < firstRow || lastColumn < firstColumn)
            return false;
        area = new KnownArea(string.Empty, range, firstRow, lastRow, firstColumn, lastColumn, "Merge");
        return true;
    }

    /// <summary>
    /// 编译期已知的实体布局矩形占用。
    /// </summary>
    private sealed class KnownArea
    {
        /// <summary>
        /// 初始化一个 <see cref="KnownArea" /> 类型的实例。
        /// </summary>
        internal KnownArea(string sheet, string address, int firstRow, int lastRow,
            int firstColumn, int lastColumn, string kind)
        {
            Sheet = sheet;
            Address = address;
            FirstRow = firstRow;
            LastRow = lastRow;
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
            Kind = kind;
        }

        /// <summary>
        /// 获取工作表名称。
        /// </summary>
        internal string Sheet { get; }
        /// <summary>
        /// 获取声明的地址。
        /// </summary>
        internal string Address { get; }
        /// <summary>
        /// 获取一基起始行。
        /// </summary>
        internal int FirstRow { get; }
        /// <summary>
        /// 获取一基末行。
        /// </summary>
        internal int LastRow { get; }
        /// <summary>
        /// 获取一基起始列。
        /// </summary>
        internal int FirstColumn { get; }
        /// <summary>
        /// 获取一基末列。
        /// </summary>
        internal int LastColumn { get; }
        /// <summary>
        /// 获取区域类型。
        /// </summary>
        internal string Kind { get; }

        /// <summary>
        /// 判断指定行列是否属于区域。
        /// </summary>
        internal bool Contains(int row, int column) => row >= FirstRow && row <= LastRow
            && column >= FirstColumn && column <= LastColumn;
    }

    /// <summary>
    /// 检查 Fluent 固定单元格的常量地址。
    /// </summary>
    private static void AnalyzeFixedCell(SyntaxNodeAnalysisContext context, InvocationExpressionSyntax current,
        IInvocationOperation operation)
    {
        var sheetArgument = GetArgument(operation, "sheetName");
        var addressArgument = GetArgument(operation, "address");
        if (sheetArgument == null || addressArgument == null)
            return;
        var sheet = ConstantString(context.SemanticModel, sheetArgument.Expression);
        var address = ConstantString(context.SemanticModel, addressArgument.Expression);
        if (address == null)
            return;
        var location = addressArgument.GetLocation();
        if (!TryNormalizeAddress(address, out var normalized))
        {
            context.ReportDiagnostic(Diagnostic.Create(InvalidAddress, location, address));
            return;
        }
        if (sheet == null)
            return;
        foreach (var previous in ReceiverChain(current))
        {
            var previousOperation = context.SemanticModel.GetOperation(previous, context.CancellationToken)
                as IInvocationOperation;
            if (previousOperation?.TargetMethod.Name != "Cell"
                || previousOperation.TargetMethod.ContainingType.OriginalDefinition.ToDisplayString() != LayoutBuilderName)
                continue;
            var previousSheetArgument = GetArgument(previousOperation, "sheetName");
            var previousAddressArgument = GetArgument(previousOperation, "address");
            if (previousSheetArgument == null || previousAddressArgument == null)
                continue;
            var previousSheet = ConstantString(context.SemanticModel, previousSheetArgument.Expression);
            var previousAddress = ConstantString(context.SemanticModel, previousAddressArgument.Expression);
            if (string.Equals(sheet, previousSheet, StringComparison.OrdinalIgnoreCase)
                && previousAddress != null && TryNormalizeAddress(previousAddress, out var prior)
                && normalized == prior)
            {
                context.ReportDiagnostic(Diagnostic.Create(DuplicateAddress, location, sheet, address));
                break;
            }
        }
    }

    /// <summary>
    /// 检查同一调用链中的动态列名称。
    /// </summary>
    private static void AnalyzeDynamicGroup(SyntaxNodeAnalysisContext context, InvocationExpressionSyntax current,
        IInvocationOperation operation)
    {
        var previousNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var previousGroups = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var previous in ReceiverChain(current))
        {
            var previousOperation = context.SemanticModel.GetOperation(previous, context.CancellationToken)
                as IInvocationOperation;
            if (previousOperation?.TargetMethod.Name != "DynamicColumnGroup"
                || previousOperation.TargetMethod.ContainingType.OriginalDefinition.ToDisplayString() != RegionBuilderName)
                continue;
            CollectGroup(context.SemanticModel, previousOperation, previousGroups, previousNames, null);
        }
        CollectGroup(context.SemanticModel, operation, previousGroups, previousNames,
            (name, location) => context.ReportDiagnostic(Diagnostic.Create(DuplicateDynamicName, location, name)));
    }

    /// <summary>
    /// 收集编译期可确定的组标识与动态列名称。
    /// </summary>
    private static void CollectGroup(SemanticModel model, IInvocationOperation invocation,
        HashSet<string> groups, HashSet<string> names, Action<string, Location>? report)
    {
        var groupArgument = GetArgument(invocation, "groupKey");
        var definitionsArgument = GetArgument(invocation, "definitions");
        if (groupArgument == null || definitionsArgument == null)
            return;
        var group = ConstantString(model, groupArgument.Expression)?.Trim();
        if (group != null && group.Length > 0 && !groups.Add(group))
            report?.Invoke(group, groupArgument.GetLocation());
        foreach (var expression in InlineArrayValues(definitionsArgument.Expression))
        {
            if (expression is not BaseObjectCreationExpressionSyntax creation)
                continue;
            if (model.GetTypeInfo(creation).Type?.ToDisplayString() != DynamicDefinitionName
                || creation.Initializer == null)
                continue;
            foreach (var assignment in creation.Initializer.Expressions.OfType<AssignmentExpressionSyntax>())
            {
                var property = model.GetSymbolInfo(assignment.Left).Symbol as IPropertySymbol;
                if (property?.ContainingType.ToDisplayString() != DynamicDefinitionName)
                    continue;
                if (property.Name == "Key" || property.Name == "Title")
                    AddName(ConstantString(model, assignment.Right), assignment.Right.GetLocation(), names, report);
                else if (property.Name == "Aliases")
                    foreach (var alias in InlineArrayValues(assignment.Right))
                        AddName(ConstantString(model, alias), alias.GetLocation(), names, report);
            }
        }
    }

    /// <summary>
    /// 按绑定的形参名称获取实参。
    /// </summary>
    private static ArgumentSyntax? GetArgument(IInvocationOperation invocation, string parameterName) =>
        invocation.Arguments.FirstOrDefault(argument => argument.Parameter?.Name == parameterName
            && !argument.IsImplicit)?.Syntax as ArgumentSyntax;

    /// <summary>
    /// 枚举单个内联数组中的元素，不混合条件表达式分支。
    /// </summary>
    private static IEnumerable<ExpressionSyntax> InlineArrayValues(ExpressionSyntax expression)
    {
        while (expression is ParenthesizedExpressionSyntax parenthesized)
            expression = parenthesized.Expression;
        return expression switch
        {
            ArrayCreationExpressionSyntax array when array.Initializer != null => array.Initializer.Expressions,
            ImplicitArrayCreationExpressionSyntax array => array.Initializer.Expressions,
            _ => Enumerable.Empty<ExpressionSyntax>()
        };
    }

    /// <summary>
    /// 记录非空名称，并报告当前集合中的重复值。
    /// </summary>
    private static void AddName(string? name, Location location, HashSet<string> names,
        Action<string, Location>? report)
    {
        name = name?.Trim();
        if (name != null && name.Length > 0 && !names.Add(name))
            report?.Invoke(name, location);
    }

    /// <summary>
    /// 枚举当前 Fluent 调用的前序调用。
    /// </summary>
    private static IEnumerable<InvocationExpressionSyntax> ReceiverChain(InvocationExpressionSyntax current)
    {
        var receiver = (current.Expression as MemberAccessExpressionSyntax)?.Expression;
        while (receiver is InvocationExpressionSyntax invocation)
        {
            yield return invocation;
            receiver = (invocation.Expression as MemberAccessExpressionSyntax)?.Expression;
        }
    }

    /// <summary>
    /// 读取编译期字符串常量。
    /// </summary>
    private static string? ConstantString(SemanticModel model, ExpressionSyntax expression)
    {
        var constant = model.GetConstantValue(expression);
        return constant.HasValue ? constant.Value as string : null;
    }

    /// <summary>
    /// 验证 A1 地址并规范化大小写与绝对引用符号。
    /// </summary>
    private static bool TryNormalizeAddress(string address, out string normalized)
    {
        normalized = string.Empty;
        var text = address.Trim().Replace("$", string.Empty).ToUpperInvariant();
        var split = 0;
        while (split < text.Length && text[split] >= 'A' && text[split] <= 'Z')
            split++;
        if (split == 0 || split == text.Length || !int.TryParse(text.Substring(split), NumberStyles.None,
                CultureInfo.InvariantCulture, out var row) || row < 1 || row > 1048576)
            return false;
        var column = 0;
        foreach (var letter in text.Substring(0, split))
        {
            var next = letter - 'A' + 1;
            if (column > (16384 - next) / 26)
                return false;
            column = column * 26 + next;
        }
        if (column < 1 || column > 16384)
            return false;
        normalized = column.ToString(CultureInfo.InvariantCulture) + ":"
            + row.ToString(CultureInfo.InvariantCulture);
        return true;
    }
}
