using System;
using Bing.Offices.Attributes;
using Bing.Offices.Validations;

namespace Bing.Offices.Mappings;

internal sealed class ExcelValidationBinding : IExcelValidationBinding
{
    private readonly FilterAttributeBase _attribute;
    private readonly IExcelValidationRule _attributeRule;
    private readonly INamedExcelValidationRule _namedRule;

    private ExcelValidationBinding(FilterAttributeBase attribute, IExcelValidationRule attributeRule)
    {
        _attribute = attribute;
        _attributeRule = attributeRule;
        Kind = ResolveKind(attribute);
        IsRaw = attribute is ExcelRequiredAttribute || attribute is ExcelRegexAttribute;
        ErrorMessage = attribute.ErrorMsg;
    }

    private ExcelValidationBinding(INamedExcelValidationRule namedRule)
    {
        _namedRule = namedRule ?? throw new ArgumentNullException(nameof(namedRule));
        Kind = ExcelValidationBindingKind.Custom;
        ErrorMessage = namedRule.ErrorMessage;
    }

    internal static ExcelValidationBinding Attribute(FilterAttributeBase attribute, IExcelValidationRule rule)
    {
        if (attribute == null)
            throw new ArgumentNullException(nameof(attribute));
        if (rule == null)
            throw new ArgumentNullException(nameof(rule));
        return new ExcelValidationBinding(attribute, rule);
    }

    internal static ExcelValidationBinding Named(INamedExcelValidationRule rule) =>
        new ExcelValidationBinding(rule);

    public ExcelValidationBindingKind Kind { get; }
    public bool IsRaw { get; }
    public string ErrorMessage { get; }

    public bool Validate(ExcelValidationContext context) => _attribute != null
        ? _attributeRule.Validate(_attribute, context)
        : _namedRule.Validate(context);

    private static ExcelValidationBindingKind ResolveKind(FilterAttributeBase attribute) => attribute switch
    {
        ExcelRequiredAttribute => ExcelValidationBindingKind.Required,
        ExcelRegexAttribute => ExcelValidationBindingKind.Regex,
        ExcelDateAttribute => ExcelValidationBindingKind.Date,
        ExcelMaxValueAttribute => ExcelValidationBindingKind.MaxValue,
        ExcelRangeAttribute => ExcelValidationBindingKind.Range,
        ExcelMaxLengthAttribute => ExcelValidationBindingKind.MaxLength,
        ExcelUniqueAttribute => ExcelValidationBindingKind.Unique,
        _ => ExcelValidationBindingKind.Custom
    };
}
