using System.Linq.Expressions;
using System.Reflection;

namespace Bing.Offices.Extensions;

/// <summary>
/// Lambda表达式扩展
/// </summary>
public static class ExpressionExtension
{
    /// <summary>
    /// 获取成员信息
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <typeparam name="TMember">成员类型</typeparam>
    /// <param name="expression">表达式</param>
    public static MemberInfo GetMemberInfo<TEntity, TMember>(this Expression<Func<TEntity, TMember>> expression)
    {
        if (expression.NodeType != ExpressionType.Lambda)
        {
            throw new ArgumentException(
                string.Format(Resources.PropertyExpression_Must_LambdaExpression, nameof(expression)),
                nameof(expression));
        }

        var lambda = expression as LambdaExpression;

        var memberExpression = ExtractMemberExpression(lambda.Body);
        if (memberExpression == null)
        {
            throw new ArgumentException(
                string.Format(Resources.PropertyExpression_Must_LambdaExpression, nameof(memberExpression)),
                nameof(memberExpression));
        }

        return memberExpression.Member;
    }

    /// <summary>
    /// 提取成员表达式
    /// </summary>
    /// <param name="expression">表达式</param>
    private static MemberExpression ExtractMemberExpression(Expression expression)
    {
        if (expression.NodeType == ExpressionType.MemberAccess)
        {
            return expression as MemberExpression;
        }

        if (expression.NodeType == ExpressionType.Convert)
        {
            var operant = ((UnaryExpression) expression).Operand;
            return ExtractMemberExpression(operant);
        }

        return null;
    }
}