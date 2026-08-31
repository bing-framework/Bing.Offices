using System.Globalization;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using Bing.Offices.Metadata;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using Bing.Offices.Npoi.Extensions;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// 负责将 NPOI 数据行转换为实体，并执行配置校验、唯一值暂存和图片值物化。
/// </summary>
internal sealed class NpoiImportRowMaterializer
{
    /// <summary>按顺序重写单元格文本的旧版转换器集合。</summary>
    private readonly IReadOnlyList<ICellValueConverter> _legacyValueConverters;

    /// <summary>
    /// 初始化行物化器。
    /// </summary>
    /// <param name="legacyValueConverters">旧版文本单元格转换器集合。</param>
    internal NpoiImportRowMaterializer(IReadOnlyList<ICellValueConverter> legacyValueConverters)
    {
        _legacyValueConverters = legacyValueConverters ?? Array.Empty<ICellValueConverter>();
    }

    /// <summary>
    /// 判断数据行是否没有任何非空单元格或图片。
    /// </summary>
    /// <param name="row">待检查的数据行。</param>
    /// <param name="bodyWhitespace">正文单元格文本的空白处理策略。</param>
    /// <param name="imageRows">包含至少一张图片的零基行索引集合。</param>
    /// <param name="rowIndex">当前行的零基索引。</param>
    /// <returns>行不存在或不包含文本和图片数据时为 true。</returns>
    internal static bool IsEmpty(IRow row, ExcelWhitespacePolicy bodyWhitespace,
        ISet<int> imageRows, int rowIndex)
    {
        return row == null
            || (imageRows == null || !imageRows.Contains(rowIndex))
            && row.Cells.All(cell => string.IsNullOrWhiteSpace(
                NpoiExcelImporter.NormalizeText(NpoiExcelImporter.GetRawStringValue(cell), bodyWhitespace)));
    }

    /// <summary>
    /// 判断列是否为图片列。
    /// </summary>
    /// <param name="column">待检查的列执行计划。</param>
    /// <returns>目标属性可接收单张或多张图片时为 true。</returns>
    internal static bool IsImageColumn(ExcelColumnPlan column) => IsImageType(column.ValueType);

    /// <summary>
    /// 构建按行列坐标索引的图片集合，并应用图片资源限制。
    /// </summary>
    /// <param name="sheet">包含图片的工作表。</param>
    /// <param name="resources">验证图片数量和字节数的资源跟踪器。</param>
    /// <param name="cancellationToken">遍历图片时检查的取消令牌。</param>
    /// <param name="imageRows">返回包含至少一张图片的零基行索引集合。</param>
    /// <returns>按零基行列坐标索引的图片集合。</returns>
    internal static IReadOnlyDictionary<(int Row, int Column), IReadOnlyList<PictureInfo>> BuildImageIndex(
        ISheet sheet, ExcelImageResourceTracker resources, CancellationToken cancellationToken,
        out HashSet<int> imageRows)
    {
        var result = new Dictionary<(int Row, int Column), IReadOnlyList<PictureInfo>>();
        imageRows = new HashSet<int>();
        foreach (var picture in sheet.GetAllPictureInfos())
        {
            cancellationToken.ThrowIfCancellationRequested();
            resources.Consume(picture.PictureData.LongLength);
            var key = (picture.MinRow, picture.MinCol);
            imageRows.Add(key.Item1);
            if (!result.TryGetValue(key, out var pictures))
                result[key] = pictures = new List<PictureInfo>();
            ((List<PictureInfo>)pictures).Add(picture);
        }
        return result;
    }

    /// <summary>
    /// 校验单行的原始单元格值。
    /// </summary>
    /// <param name="row">待校验的数据行。</param>
    /// <param name="columns">按零基列索引排列的列执行计划。</param>
    /// <param name="duplicateValues">兼容旧校验规则的重复值状态。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="rowIndex">当前行的零基索引。</param>
    /// <param name="validateMode">发生校验失败后的继续策略。</param>
    /// <param name="culture">校验上下文使用的区域性。</param>
    /// <param name="bodyWhitespace">正文单元格文本的空白处理策略。</param>
    /// <param name="errors">接收原始值校验错误的收集器。</param>
    /// <returns>当前行全部原始值校验通过时为 true。</returns>
    internal bool ValidateRawValues(IRow row, IReadOnlyDictionary<int, ExcelColumnPlan> columns,
        IDictionary<string, HashSet<string>> duplicateValues, string sheetName, int rowIndex,
        ValidateMode validateMode, CultureInfo culture, ExcelWhitespacePolicy bodyWhitespace,
        ExcelImportErrorCollector errors)
    {
        var valid = true;
        foreach (var column in columns)
        {
            var cell = row.GetCell(column.Key);
            var cellValue = NormalizeCellValue(ApplyLegacyTextConverters(cell,
                    NpoiExcelImporter.ReadCellValue(cell)),
                column.Value.Property.ImportWhitespace ?? bodyWhitespace);
            var value = cellValue.Text;
            foreach (var binding in column.Value.ValidationBindings.Where(binding => binding.IsRaw))
            {
                var context = new ExcelValidationContext(value, sheetName, rowIndex + 1, column.Key + 1,
                    column.Value.Property.Name, null, column.Value.ValueType, cellValue, culture);
                bool isValid;
                try
                {
                    isValid = binding.Validate(context);
                }
                catch (Exception exception)
                {
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.Validation, exception.Message, sheetName,
                        rowIndex + 1, column.Key + 1, column.Value.Property.Name, GetErrorColumnKey(column.Value),
                        column.Value.HeaderName, cellValue.Value ?? cellValue.Text));
                    valid = false;
                    if (validateMode == ValidateMode.StopOnFirstFailure)
                        return false;
                    continue;
                }
                if (isValid)
                    continue;
                errors.Add(new ExcelImportError(GetValidationErrorCode(binding), binding.ErrorMessage, sheetName,
                    rowIndex + 1, column.Key + 1, column.Value.Property.Name, GetErrorColumnKey(column.Value),
                    column.Value.HeaderName, cellValue.Value ?? cellValue.Text));
                valid = false;
                if (validateMode == ValidateMode.StopOnFirstFailure)
                    return false;
            }
        }
        return valid;
    }

    /// <summary>
    /// 将单行数据转换为实体，并执行转换后校验和唯一值校验。
    /// </summary>
    /// <typeparam name="T">要物化的实体类型。</typeparam>
    /// <param name="row">待转换的数据行。</param>
    /// <param name="columns">按零基列索引排列的列执行计划。</param>
    /// <param name="duplicateValues">兼容旧校验规则的重复值状态。</param>
    /// <param name="uniqueTracker">负责当前行唯一值预留、提交和回滚的跟踪器。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="rowIndex">当前行的零基索引。</param>
    /// <param name="validateMode">发生校验失败后的继续策略。</param>
    /// <param name="configuredValidationEnabled">是否执行配置校验规则。</param>
    /// <param name="errors">接收转换和校验错误的收集器。</param>
    /// <param name="culture">文本转换和校验使用的区域性。</param>
    /// <param name="bodyWhitespace">正文单元格文本的空白处理策略。</param>
    /// <param name="dynamicTargetGetter">从实体取得动态值目标字典的委托。</param>
    /// <param name="imageIndex">按行列坐标索引的图片集合。</param>
    /// <param name="item">成功时返回已物化的实体；失败时为 null。</param>
    /// <returns>当前行成功转换并通过校验时为 true。</returns>
    internal bool TryCreateItem<T>(IRow row, IReadOnlyDictionary<int, ExcelColumnPlan> columns,
        IDictionary<string, HashSet<string>> duplicateValues, UniqueTracker uniqueTracker,
        string sheetName, int rowIndex, ValidateMode validateMode, bool configuredValidationEnabled,
        ExcelImportErrorCollector errors, CultureInfo culture, ExcelWhitespacePolicy bodyWhitespace,
        Func<object, object> dynamicTargetGetter,
        IReadOnlyDictionary<(int Row, int Column), IReadOnlyList<PictureInfo>> imageIndex, out T item)
        where T : class, new()
    {
        item = new T();
        Dictionary<string, object> dynamicValues = null;
        foreach (var column in columns)
        {
            var cellValue = default(ExcelCellValue);
            try
            {
                var cell = row.GetCell(column.Key);
                var images = imageIndex == null ? null : FindImages(imageIndex, rowIndex, column.Key);
                var imageMultiplicity = column.Value.ImageMultiplicity;
                if (images != null && images.Count > 1
                    && imageMultiplicity == ExcelImageMultiplicityPolicy.Fail)
                {
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidInput, "同一单元格存在多个图片。",
                        sheetName, rowIndex + 1, column.Key + 1, column.Value.Property.Name,
                        GetErrorColumnKey(column.Value)));
                    item = null;
                    return false;
                }
                var image = images?.FirstOrDefault();
                cellValue = image == null
                    ? NormalizeCellValue(ApplyLegacyTextConverters(cell, NpoiExcelImporter.ReadCellValue(cell)),
                        column.Value.Property.ImportWhitespace ?? bodyWhitespace)
                    : new ExcelCellValue(image, string.Empty, ExcelCellKind.Empty);
                var value = cellValue.Text;
                if (column.Value.IsDynamic)
                {
                    dynamicValues ??= new Dictionary<string, object>(StringComparer.Ordinal);
                    object dynamicConvertedValue;
                    if (column.Value.DynamicDefinition == null)
                        dynamicConvertedValue = value;
                    else
                    {
                        var propertyType = column.Value.ValueType;
                        if (image != null && IsImageType(propertyType))
                            dynamicConvertedValue = ConvertImages(images, propertyType, imageMultiplicity);
                        else
                            dynamicConvertedValue = column.Value.ConvertFrom(value, cellValue, sheetName,
                                rowIndex + 1, column.Key + 1, culture);
                    }
                    if (configuredValidationEnabled && !ValidateColumnValue(value, cellValue,
                        dynamicConvertedValue, column.Value, duplicateValues, uniqueTracker, sheetName, rowIndex,
                        validateMode, culture, errors))
                    {
                        item = null;
                        return false;
                    }
                    dynamicValues[column.Value.Key] = dynamicConvertedValue;
                    continue;
                }
                var converted = image != null && IsImageType(column.Value.ValueType)
                    ? ConvertImages(images, column.Value.ValueType, imageMultiplicity)
                    : column.Value.ConvertFrom(value, cellValue, sheetName, rowIndex + 1, column.Key + 1, culture);
                if (configuredValidationEnabled && !ValidateColumnValue(value, cellValue, converted,
                    column.Value, duplicateValues, uniqueTracker, sheetName, rowIndex, validateMode, culture,
                    errors))
                {
                    item = null;
                    return false;
                }
                column.Value.Setter(item, converted);
            }
            catch (Exception exception)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.ValueConversion, exception.Message, sheetName,
                    rowIndex + 1, column.Key + 1, column.Value.Property.Name, GetErrorColumnKey(column.Value),
                    column.Value.HeaderName, cellValue?.Value ?? cellValue?.Text));
                item = null;
                return false;
            }
        }
        if (dynamicValues != null)
        {
            var target = dynamicTargetGetter?.Invoke(item) as IDictionary<string, object>;
            if (target != null)
            {
                foreach (var pair in dynamicValues)
                    target[pair.Key] = pair.Value;
            }
            else
                columns.Values.First(column => column.IsDynamic).Setter(item, dynamicValues);
        }
        return true;
    }

    /// <summary>
    /// 使用旧版文本转换器读取单元格文本，同时保留 typed cell 元数据。
    /// </summary>
    private ExcelCellValue ApplyLegacyTextConverters(ICell cell, ExcelCellValue cellValue)
    {
        if (cell == null || _legacyValueConverters.Count == 0)
            return cellValue;
        var text = cellValue.Text;
        foreach (var converter in _legacyValueConverters)
        {
            var converted = converter.GetStringValue(cell);
            if (converted != null)
                text = converted;
        }
        return text == cellValue.Text
            ? cellValue
            : new ExcelCellValue(cellValue.Value, text, cellValue.Kind, cellValue.CachedKind,
                cellValue.Formula, cellValue.ErrorCode, cellValue.FormatIndex);
    }

    /// <summary>
    /// 按列配置校验转换后的值和唯一性。
    /// </summary>
    private static bool ValidateColumnValue(string value, ExcelCellValue cellValue, object convertedValue,
        ExcelColumnPlan column, IDictionary<string, HashSet<string>> duplicateValues,
        UniqueTracker uniqueTracker, string sheetName, int rowIndex, ValidateMode validateMode,
        CultureInfo culture, ExcelImportErrorCollector errors)
    {
        var valid = true;
        var property = column.Property;
        var context = new ExcelValidationContext(value, sheetName, rowIndex + 1, column.ColumnIndex + 1,
            column.IsDynamic ? column.Key : property.Name, convertedValue, column.ValueType, cellValue, culture);
        foreach (var binding in column.ValidationBindings.Where(binding => !binding.IsRaw
                     && binding.Kind != ExcelValidationBindingKind.Unique))
        {
            bool isValid;
            try
            {
                isValid = binding.Validate(context);
            }
            catch (Exception exception)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.Validation, exception.Message, sheetName,
                    rowIndex + 1, column.ColumnIndex + 1, property.Name, GetErrorColumnKey(column),
                    column.HeaderName, cellValue.Value ?? cellValue.Text));
                valid = false;
                if (validateMode == ValidateMode.StopOnFirstFailure)
                    return false;
                continue;
            }
            if (isValid)
                continue;
            errors.Add(new ExcelImportError(GetValidationErrorCode(binding), binding.ErrorMessage, sheetName,
                rowIndex + 1, column.ColumnIndex + 1, property.Name, GetErrorColumnKey(column),
                column.HeaderName, cellValue.Value ?? cellValue.Text));
            valid = false;
            if (validateMode == ValidateMode.StopOnFirstFailure)
                return false;
        }
        if (column.IsUnique)
        {
            bool reserved;
            try
            {
                reserved = uniqueTracker.TryReserve(column.Key, value, false, column.UniqueIgnoreEmpty,
                    rowIndex + 1);
            }
            catch (Exception exception)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.ResourceLimit, exception.Message, sheetName,
                    rowIndex + 1, column.ColumnIndex + 1, property.Name, GetErrorColumnKey(column),
                    column.HeaderName, cellValue.Value ?? cellValue.Text));
                return false;
            }
            if (!reserved)
            {
                var firstRowNumber = uniqueTracker.TryGetFirstRowNumber(column.Key, value, out var firstRow)
                    ? firstRow
                    : (int?)null;
                errors.Add(new ExcelImportError(ExcelImportErrorCode.Validation, "重复数据", sheetName,
                    rowIndex + 1, column.ColumnIndex + 1, property.Name, GetErrorColumnKey(column),
                    column.HeaderName, cellValue.Value ?? cellValue.Text, firstRowNumber));
                return false;
            }
        }
        return valid;
    }

    /// <summary>
    /// 规范化单元格值中的正文空白。
    /// </summary>
    private static ExcelCellValue NormalizeCellValue(ExcelCellValue cellValue, ExcelWhitespacePolicy policy)
    {
        if (cellValue == null || cellValue.Kind != ExcelCellKind.Text && cellValue.Kind != ExcelCellKind.Formula)
            return cellValue;
        var text = NpoiExcelImporter.NormalizeText(cellValue.Text, policy);
        return text == cellValue.Text
            ? cellValue
            : new ExcelCellValue(cellValue.Value, text, cellValue.Kind, cellValue.CachedKind,
                cellValue.Formula, cellValue.ErrorCode, cellValue.FormatIndex);
    }

    /// <summary>
    /// 查找指定行列位置的图片。
    /// </summary>
    private static IReadOnlyList<PictureInfo> FindImages(
        IReadOnlyDictionary<(int Row, int Column), IReadOnlyList<PictureInfo>> imageIndex, int row, int column)
    {
        return imageIndex.TryGetValue((row, column), out var pictures) ? pictures : null;
    }

    /// <summary>
    /// 判断目标类型是否为图片或图片集合。
    /// </summary>
    private static bool IsImageType(Type type)
    {
        if (type == null || type == typeof(byte[]) || type == typeof(ExcelImageData))
            return type != null;
        if (type.IsArray)
            return type.GetElementType() == typeof(byte[]) || type.GetElementType() == typeof(ExcelImageData);
        if (type.IsGenericType && type.GetGenericArguments().Length == 1)
        {
            var elementType = type.GetGenericArguments()[0];
            if (elementType == typeof(byte[]) || elementType == typeof(ExcelImageData))
                return true;
        }
        return GetImageElementType(type) != null;
    }

    /// <summary>
    /// 将图片集合转换为目标属性类型。
    /// </summary>
    private static object ConvertImages(IReadOnlyList<PictureInfo> pictures, Type targetType,
        ExcelImageMultiplicityPolicy policy)
    {
        if (policy == ExcelImageMultiplicityPolicy.All)
        {
            var elementType = GetImageElementType(targetType);
            if (elementType == typeof(byte[]))
            {
                var values = pictures.Select(picture => picture.PictureData).ToArray();
                return targetType.IsArray ? values : values.ToList();
            }
            if (elementType == typeof(ExcelImageData))
            {
                var values = pictures.Select(picture => (ExcelImageData)ConvertImage(picture,
                    typeof(ExcelImageData))).ToList();
                return targetType.IsArray ? values.ToArray() : values;
            }
        }
        return ConvertImage(pictures[0], targetType);
    }

    /// <summary>
    /// 解析图片集合元素类型。
    /// </summary>
    private static Type GetImageElementType(Type type)
    {
        if (type == null)
            return null;
        if (type.IsArray)
            return type.GetElementType();
        if (type.IsGenericType && type.GetGenericArguments().Length == 1)
        {
            var directElementType = type.GetGenericArguments()[0];
            if (directElementType == typeof(byte[]) || directElementType == typeof(ExcelImageData))
                return directElementType;
        }
        return type.GetInterfaces().Concat(new[] { type })
            .Where(candidate => candidate.IsGenericType
                && candidate.GetGenericTypeDefinition() == typeof(IEnumerable<>))
            .Select(candidate => candidate.GetGenericArguments()[0])
            .FirstOrDefault(candidate => candidate == typeof(byte[]) || candidate == typeof(ExcelImageData));
    }

    /// <summary>
    /// 将单张图片转换为目标属性类型。
    /// </summary>
    private static object ConvertImage(PictureInfo picture, Type targetType)
    {
        if (targetType == typeof(byte[]))
            return picture.PictureData;
        return new ExcelImageData(picture.PictureData, ResolveImageContentType(picture.PictureData),
            picture.MinRow + 1, picture.MinCol + 1, picture.MaxRow + 1, picture.MaxCol + 1);
    }

    /// <summary>
    /// 根据图片头部识别常见图片内容类型。
    /// </summary>
    private static string ResolveImageContentType(byte[] bytes)
    {
        if (bytes?.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
            return "image/jpeg";
        if (bytes?.Length >= 6 && bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46)
            return "image/gif";
        return "image/png";
    }

    /// <summary>
    /// 映射验证绑定到公共导入错误码。
    /// </summary>
    private static ExcelImportErrorCode GetValidationErrorCode(IExcelValidationBinding binding) =>
        binding.Kind == ExcelValidationBindingKind.MaxLength
            ? ExcelImportErrorCode.MaxLength
            : binding.Kind == ExcelValidationBindingKind.MaxValue
                ? ExcelImportErrorCode.MaxValue
                : ExcelImportErrorCode.Validation;

    /// <summary>
    /// 获取固定列或动态列的错误定位键。
    /// </summary>
    private static string GetErrorColumnKey(ExcelColumnPlan column) =>
        column.IsDynamic ? column.Key : column.Property.Name;
}
