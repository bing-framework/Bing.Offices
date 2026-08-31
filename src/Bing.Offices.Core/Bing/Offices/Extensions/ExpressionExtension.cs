using System.Linq.Expressions;
using System.Reflection;

namespace Bing.Offices.Extensions;

/// <summary>
    /// 从 Lambda 表达式中解析实体成员信息的扩展。
/// </summary>
public static class ExpressionExtension
{
    /// <summary>
    /// 解析表达式主体所指向的属性或字段成员。
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <typeparam name="TMember">成员类型</typeparam>
    /// <param name="expression">表达式</param>
    /// <returns>表达式指向的成员信息。</returns>
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
    /// 从成员访问或装箱转换表达式中提取成员访问节点。
    /// </summary>
    /// <param name="expression">表达式</param>
    /// <returns>找到的成员表达式；表达式不代表成员访问时返回 null。</returns>
    private static MemberExpression ExtractMemberExpression(Expression expression)
    {
        if (expression.NodeType == ExpressionType.MemberAccess)
        {
            return expression as MemberExpression;
        }

        if (expression.NodeType == ExpressionType.Convert)
        {
            var operant = ((UnaryExpression)expression).Operand;
            return ExtractMemberExpression(operant);
        }

        return null;
    }
}
