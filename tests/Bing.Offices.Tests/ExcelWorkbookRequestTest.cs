using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Npoi.Exports;
using Bing.Offices.Npoi.Imports;
using Bing.Offices.Npoi;
using Bing.Offices.Styles;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// Workbook Request、动态列和关系绑定测试。
/// </summary>
public sealed class ExcelWorkbookRequestTest
{
    private static readonly IReadOnlyList<ExcelDynamicColumnDefinition> TenantColumns =
        new[]
        {
            new ExcelDynamicColumnDefinition
            {
                Key = "region",
                Title = "地区",
                Aliases = new[] { "旧地区" },
                DataType = typeof(string),
                Order = 10
            },
            new ExcelDynamicColumnDefinition
            {
                Key = "amount",
                Title = "金额",
                DataType = typeof(decimal),
                NumberFormat = "0.00",
                Order = 20
            }
        };

    /// <summary>
    /// 测试 - Workbook 导出请求应使用同一 Cell Writer 写入多个 Sheet，并按动态 Key 读取字典值。
    /// </summary>
    [Fact]
    public void Export_WorkbookRequest_ShouldWriteMultipleSheetsAndTypedDynamicKeys()
    {
        // Arrange
        var request = ExcelExport.Workbook(workbook =>
        {
            workbook.AddSheet("订单", new[]
            {
                new ExportOrder { OrderNo = "O-1", CustomFields = new Dictionary<string, object>
                {
                    ["region"] = "华东", ["amount"] = 12.5m
                } }
            }, sheet => sheet
                .HeaderRowIndex(2)
                .DynamicColumns(order => order.CustomFields, TenantColumns));
            workbook.AddSheet("客户", new[] { new ExportCustomer { Name = "客户 A" } });
        });
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        Assert.Equal(2, workbook.NumberOfSheets);
        var orderSheet = workbook.GetSheet("订单");
        Assert.Equal("地区", orderSheet.GetRow(2).GetCell(1).StringCellValue);
        Assert.Equal("金额", orderSheet.GetRow(2).GetCell(2).StringCellValue);
        Assert.Equal("华东", orderSheet.GetRow(3).GetCell(1).StringCellValue);
        Assert.Equal(CellType.Numeric, orderSheet.GetRow(3).GetCell(2).CellType);
        Assert.Equal("客户 A", workbook.GetSheet("客户").GetRow(1).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 测试 - 异构 Workbook 导入应按每个 Sheet 的 HeaderRowIndex/DataRowStartIndex 解析别名和动态类型。
    /// </summary>
    [Fact]
    public void Import_WorkbookRequest_ShouldUsePerSheetHeaderAndDynamicDefinitions()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var orders = workbook.CreateSheet("订单");
            orders.CreateRow(2).CreateCell(0).SetCellValue("订单号");
            orders.GetRow(2).CreateCell(1).SetCellValue("旧地区");
            orders.GetRow(2).CreateCell(2).SetCellValue("金额");
            orders.CreateRow(3).CreateCell(0).SetCellValue("O-1");
            orders.GetRow(3).CreateCell(1).SetCellValue("华东");
            orders.GetRow(3).CreateCell(2).SetCellValue("12.5");
            var details = workbook.CreateSheet("明细");
            details.CreateRow(1).CreateCell(0).SetCellValue("订单号");
            details.GetRow(1).CreateCell(1).SetCellValue("名称");
            details.CreateRow(2).CreateCell(0).SetCellValue("O-1");
            details.GetRow(2).CreateCell(1).SetCellValue("商品");
        }));
        var request = ExcelImport.Workbook<OrderWorkbook>(workbook =>
        {
            workbook.Sheet("订单", root => root.Orders, sheet => sheet
                .HeaderRowIndex(2)
                .DataRowStartIndex(3)
                .DynamicColumns(order => order.CustomFields, TenantColumns));
            workbook.Sheet("明细", root => root.Details, sheet => sheet
                .HeaderRowIndex(1)
                .DataRowStartIndex(2));
            workbook.HasMany(root => root.Orders, root => root.Details,
                order => order.OrderNo, detail => detail.OrderNo,
                order => order.DetailItems);
        });

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.True(result.IsSuccess);
        var order = Assert.Single(result.Workbook.Orders);
        Assert.Equal("华东", order.CustomFields["region"]);
        Assert.Equal(12.5m, order.CustomFields["amount"]);
        Assert.Equal("商品", Assert.Single(order.DetailItems).Name);
        Assert.Equal(2, result.Sheets.Count);
        Assert.All(result.Sheets, sheet => Assert.Empty(sheet.Errors));
    }

    /// <summary>
    /// 测试 - 泛型整型关系键应绑定成功，找不到父项时错误应保留子 Sheet、行号和 Key。
    /// </summary>
    [Fact]
    public void Import_RelationWithNonStringKey_ShouldKeepSourceLocation()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var parents = workbook.CreateSheet("Parents");
            parents.CreateRow(0).CreateCell(0).SetCellValue("Id");
            parents.CreateRow(1).CreateCell(0).SetCellValue(1);
            var children = workbook.CreateSheet("Children");
            children.CreateRow(0).CreateCell(0).SetCellValue("ParentId");
            children.CreateRow(1).CreateCell(0).SetCellValue(2);
        }));
        var request = ExcelImport.Workbook<IntRelationWorkbook>(builder =>
        {
            builder.Sheet("Parents", root => root.Parents);
            builder.Sheet("Children", root => root.Children);
            builder.HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId, parent => parent.Children);
        });

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.True(result.Errors.Count == 1,
            string.Join("; ", result.Errors.Select(error =>
                $"{error.Code}|{error.SheetName}|{error.RowIndex}|{error.ColumnKey}|{error.RawValue}")));
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Relationship, error.Code);
        Assert.Equal("Children", error.SheetName);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal("Key", error.ColumnKey);
        Assert.Equal(2, error.RawValue);
    }

    /// <summary>
    /// 测试 - 关系绑定应使用调用方提供的 comparer，大小写不同的字符串键也应关联成功。
    /// </summary>
    [Fact]
    public void Import_RelationWithCustomComparer_ShouldBindCaseInsensitiveKeys()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var parents = workbook.CreateSheet("Parents");
            parents.CreateRow(0).CreateCell(0).SetCellValue("订单号");
            parents.CreateRow(1).CreateCell(0).SetCellValue("A-1");
            var children = workbook.CreateSheet("Children");
            children.CreateRow(0).CreateCell(0).SetCellValue("订单号");
            children.GetRow(0).CreateCell(1).SetCellValue("名称");
            children.CreateRow(1).CreateCell(0).SetCellValue("a-1");
            children.GetRow(1).CreateCell(1).SetCellValue("商品");
        }));
        var request = ExcelImport.Workbook<OrderWorkbook>(builder =>
        {
            builder.Sheet("Parents", root => root.Orders);
            builder.Sheet("Children", root => root.Details);
            builder.HasMany(root => root.Orders, root => root.Details,
                parent => parent.OrderNo, child => child.OrderNo,
                parent => parent.DetailItems, StringComparer.OrdinalIgnoreCase);
        });

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.True(result.IsSuccess, string.Join("; ", result.Errors.Select(error => error.Message)));
        Assert.Equal("商品", Assert.Single(Assert.Single(result.Workbook.Orders).DetailItems).Name);
    }

    /// <summary>
    /// 测试 - fixed 与 dynamic 列应由同一不可变计划承载执行元数据和预编译访问器。
    /// </summary>
    [Fact]
    public void ColumnPlan_FixedAndDynamic_ShouldShareCompiledExecutionMetadata()
    {
        // Arrange
        var map = Bing.Offices.Mappings.ExcelTypeMapFactory.Get<ImportOrder>();
        var fixedProperty = map.Properties.Single(property => property.Name == nameof(ImportOrder.OrderNo));
        var dynamicProperty = map.Properties.Single(property => property.Name == nameof(ImportOrder.CustomFields));
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "地区",
            DataType = typeof(string),
            ConverterName = "region",
            ValidatorName = "required"
        };
        var fixedPlan = new ExcelColumnPlan(fixedProperty.Title, fixedProperty, false, 0, null);
        var dynamicPlan = new ExcelColumnPlan(definition.Title, dynamicProperty, true, 1, definition,
            definition.Key);

        // Assert
        Assert.Same(fixedProperty.Getter, fixedPlan.Getter);
        Assert.Same(fixedProperty.Setter, fixedPlan.Setter);
        Assert.Equal(typeof(string), fixedPlan.ValueType);
        Assert.Equal(typeof(string), dynamicPlan.ValueType);
        Assert.Equal("region", dynamicPlan.Key);
        Assert.Equal("region", dynamicPlan.ConverterName);
        Assert.Equal("required", dynamicPlan.ValidatorName);
        Assert.Equal(ExcelImageMultiplicityPolicy.First, dynamicPlan.ImageMultiplicity);
    }

    /// <summary>
    /// 测试 - Before/After 与物理列索引不能同时配置。
    /// </summary>
    [Fact]
    public void DynamicColumnPlacement_CombinedRelativeAndPhysicalIndex_ShouldFail()
    {
        // Arrange
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "地区",
            Placement = ExcelColumnPlacement.Before("OrderNo"),
            PhysicalColumnIndex = 1
        };

        // Act
        var action = () => ExcelExport.Workbook(workbook => workbook.AddSheet("订单",
            new[] { new ExportOrder { OrderNo = "O-1", CustomFields = new Dictionary<string, object>() } },
            sheet => sheet.DynamicColumns(order => order.CustomFields, new[] { definition })));

        // Assert
        Assert.Throws<ArgumentException>(action);
    }

    /// <summary>
    /// 测试 - 动态值应按 DataType 写为 Numeric Cell，未知 Key 在 Fail 策略下应拒绝导出。
    /// </summary>
    [Fact]
    public void DynamicColumn_DataTypeAndUnknownValuePolicy_ShouldBeEnforced()
    {
        // Arrange
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "amount",
            Title = "金额",
            DataType = typeof(decimal)
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("订单",
            new[] { new ExportOrder { OrderNo = "O-1", CustomFields = new Dictionary<string, object>
            {
                ["amount"] = "12.5", ["unknown"] = "拒绝"
            } } }, sheet => sheet
                .DynamicColumns(order => order.CustomFields, new[] { definition })
                .UnknownDynamicValues(ExcelUnknownDynamicValuePolicy.Fail)));
        using var destination = new MemoryStream();

        // Act
        var action = () => new NpoiExcelExporter().Export(request, destination);

        // Assert
        Assert.Throws<InvalidOperationException>(action);

        var validRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("订单",
            new[] { new ExportOrder { OrderNo = "O-1", CustomFields = new Dictionary<string, object>
            {
                ["amount"] = "12.5"
            } } }, sheet => sheet.DynamicColumns(order => order.CustomFields, new[] { definition })));
        new NpoiExcelExporter().Export(validRequest, destination);
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        Assert.Equal(CellType.Numeric, workbook.GetSheet("订单").GetRow(1).GetCell(1).CellType);
    }

    /// <summary>
    /// 测试 - Before 放置应改变最终物理列顺序，并保持固定列可读。
    /// </summary>
    [Fact]
    public void DynamicColumnPlacement_BeforeFixedColumn_ShouldChangePhysicalOrder()
    {
        // Arrange
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "地区",
            Placement = ExcelColumnPlacement.Before("OrderNo")
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("订单",
            new[] { new ExportOrder { OrderNo = "O-1", CustomFields = new Dictionary<string, object>
            {
                ["region"] = "华东"
            } } }, sheet => sheet.DynamicColumns(order => order.CustomFields, new[] { definition })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var sheet = workbook.GetSheet("订单");
        Assert.Equal("地区", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("OrderNo", sheet.GetRow(0).GetCell(1).StringCellValue);
        Assert.Equal("华东", sheet.GetRow(1).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 测试 - 模板命名区域应确定写入起点，保留模板其它内容，并按 leaveOpen 约定处理输入流。
    /// </summary>
    [Fact]
    public void Export_TemplateRegion_ShouldPreserveTemplateAndKeepInputOpenWhenRequested()
    {
        // Arrange
        using var template = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("订单");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("保留内容");
            var name = workbook.CreateName();
            name.NameName = "OrderTable";
            name.RefersToFormula = "'订单'!$B$3:$C$20";
            workbook.CreateSheet("说明").CreateRow(0).CreateCell(0).SetCellValue("说明");
        }));
        var request = ExcelExport.Workbook(workbook =>
        {
            workbook.UseTemplate(template, leaveOpen: true)
                .AddSheet("订单", new[] { new ExportCustomer { Name = "客户 A" } },
                    sheet => sheet.UseTemplateRegion("OrderTable"));
        });
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        Assert.True(template.CanRead);
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        Assert.Equal("保留内容", result.GetSheet("订单").GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("客户 A", result.GetSheet("订单").GetRow(3).GetCell(1).StringCellValue);
        Assert.Equal("说明", result.GetSheet("说明").GetRow(0).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 测试 - 同一 Workbook 中相同请求样式应复用样式定义，并保留 XLSX 自定义颜色。
    /// </summary>
    [Fact]
    public void Export_RequestStyle_ShouldReuseStyleAndWriteCustomXlsxColor()
    {
        // Arrange
        var style = new ExcelCellStyle
        {
            Bold = true,
            FillPattern = ExcelFillPattern.Solid,
            ForegroundColor = new ExcelColor("FF123456")
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("客户",
            new[] { new ExportCustomer { Name = "客户 A" }, new ExportCustomer { Name = "客户 B" } },
            sheet => sheet.HeaderStyle(style).BodyStyle(style)));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var sheet = workbook.GetSheet("客户");
        Assert.Equal(sheet.GetRow(0).GetCell(0).CellStyle.Index, sheet.GetRow(1).GetCell(0).CellStyle.Index);
        Assert.Equal("123456", BitConverter.ToString(((XSSFCellStyle)sheet.GetRow(0).GetCell(0).CellStyle)
            .FillForegroundXSSFColor.RGB).Replace("-", string.Empty));
    }

    /// <summary>
    /// 测试 - 请求样式应逐属性叠加到模板样式，不得覆盖模板数字格式。
    /// </summary>
    [Fact]
    public void Export_RequestStyle_ShouldPreserveTemplateNumberFormat()
    {
        // Arrange
        using var template = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("订单");
            var style = workbook.CreateCellStyle();
            style.DataFormat = workbook.CreateDataFormat().GetFormat("0.00");
            sheet.CreateRow(3).CreateCell(1).CellStyle = style;
            var name = workbook.CreateName();
            name.NameName = "OrderTable";
            name.RefersToFormula = "'订单'!$B$3:$C$20";
        }));
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("订单", new[] { new ExportCustomer { Name = "客户 A" } },
                sheet => sheet.UseTemplateRegion("OrderTable")
                    .BodyStyle(new ExcelCellStyle { Bold = true })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        var cell = result.GetSheet("订单").GetRow(3).GetCell(1);
        Assert.Equal("0.00", cell.CellStyle.GetDataFormatString());
        Assert.True(result.GetSheet("订单").GetRow(3).GetCell(1).CellStyle.Index > 0);
    }

    /// <summary>
    /// 测试 - 模板区域起点应同时偏移自定义表头的相对行列。
    /// </summary>
    [Fact]
    public void Export_TemplateRegion_ShouldOffsetCustomHeaders()
    {
        // Arrange
        using var template = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("订单");
            var name = workbook.CreateName();
            name.NameName = "OrderTable";
            name.RefersToFormula = "'订单'!$B$3:$C$20";
        }));
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("订单", new[] { new ExportCustomer { Name = "客户 A" } },
                sheet => sheet.UseTemplateRegion("OrderTable")
                    .HeaderRowIndex(1)
                    .DataRowStartIndex(2)
                    .HeaderRows(new[] { new ExcelHeaderRow(0,
                        new[] { new ExcelHeaderCell(2, "扩展表头") }) })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        Assert.Equal("扩展表头", result.GetSheet("订单").GetRow(2).GetCell(3).StringCellValue);
    }

    /// <summary>
    /// 测试 - 动态列指定命名转换器时，应在默认类型转换前使用该转换器。
    /// </summary>
    [Fact]
    public void Export_DynamicColumnNamedConverter_ShouldConvertValue()
    {
        // Arrange
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "地区",
            DataType = typeof(string),
            ConverterName = "upper-region"
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("订单",
            new[] { new ExportOrder { OrderNo = "O-1", CustomFields = new Dictionary<string, object>
            {
                ["region"] = "east"
            } } }, sheet => sheet.DynamicColumns(order => order.CustomFields, new[] { definition })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter(new[] { new UpperRegionConverter() }).Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        Assert.Equal("EAST", result.GetSheet("订单").GetRow(1).GetCell(1).StringCellValue);
    }

    /// <summary>
    /// 测试 - 按从零开始的索引选择 Sheet 时，应读取指定工作表而不是依赖名称。
    /// </summary>
    [Fact]
    public void Import_SheetSelectorByIndex_ShouldReadSelectedSheet()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            workbook.CreateSheet("忽略").CreateRow(0).CreateCell(0).SetCellValue("Code");
            var sheet = workbook.CreateSheet("目标");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("B-1");
        }));
        var request = ExcelImport.Workbook<SelectorWorkbook>(workbook =>
            workbook.Sheet(ExcelSheetSelector.ByIndex(1), root => root.Rows));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.True(result.IsSuccess);
        Assert.Equal("B-1", Assert.Single(result.Workbook.Rows).Code);
        Assert.Equal("目标", Assert.Single(result.Sheets).Name);
    }

    /// <summary>
    /// 测试 - ReadColumns 应限制表头绑定范围，并允许范围外的辅助列存在。
    /// </summary>
    [Fact]
    public void Import_ReadColumns_ShouldIgnoreColumnsOutsideRange()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("辅助列");
            sheet.GetRow(0).CreateCell(1).SetCellValue("Code");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("ignored");
            sheet.GetRow(1).CreateCell(1).SetCellValue("A-1");
        }));
        var request = ExcelImport.Workbook<SelectorWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet => sheet.ReadColumns(1, 1)));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.True(result.IsSuccess);
        Assert.Equal("A-1", Assert.Single(result.Workbook.Rows).Code);
    }

    /// <summary>
    /// 测试 - Header Trim 与 Body Preserve 应分别生效，正文原始首尾空白不得被 Reader 提前丢弃。
    /// </summary>
    [Fact]
    public void Import_WhitespacePolicies_ShouldApplyHeaderAndBodySeparately()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(" Code ");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("  A-1  ");
        }));
        var request = ExcelImport.Workbook<SelectorWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet => sheet
                .HeaderWhitespace(ExcelWhitespacePolicy.Trim)
                .BodyWhitespace(ExcelWhitespacePolicy.Preserve)));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.True(result.IsSuccess);
        Assert.Equal("  A-1  ", Assert.Single(result.Workbook.Rows).Code);
    }

    /// <summary>
    /// 测试 - 动态列应写入请求指定的动态目标，而不是映射属性的默认 Setter。
    /// </summary>
    [Fact]
    public void Import_DynamicTargetGetter_ShouldWriteRequestedDictionary()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Extra");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("value");
        }));
        var definition = new ExcelDynamicColumnDefinition { Key = "extra", Title = "Extra" };
        var request = ExcelImport.Workbook<DynamicTargetWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet => sheet
                .DynamicColumns(row => row.TargetFields, new[] { definition })));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);
        Assert.True(result.IsSuccess, string.Join("; ", result.Errors.Select(error => error.Message)));
        var item = Assert.Single(result.Workbook.Rows);

        // Assert
        Assert.Equal("value", item.TargetFields["extra"]);
        Assert.Null(item.MappedFields);
    }

    /// <summary>
    /// 测试 - 动态列导入应使用定义中的命名双向转换器。
    /// </summary>
    [Fact]
    public void Import_DynamicColumnNamedConverter_ShouldConvertValue()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("地区");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("east");
        }));
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "地区",
            ConverterName = "upper-region"
        };
        var request = ExcelImport.Workbook<DynamicTargetWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet => sheet
                .DynamicColumns(row => row.TargetFields, new[] { definition })));

        // Act
        var result = new NpoiExcelImporter(valueConverters: new IExcelValueConverter[] { new UpperRegionConverter() })
            .Import(source, request);

        // Assert
        Assert.True(result.IsSuccess, string.Join("; ", result.Errors.Select(error => error.Message)));
        Assert.Equal("EAST", Assert.Single(result.Workbook.Rows).TargetFields["region"]);
    }

    /// <summary>
    /// 测试 - 固定列转换器应在列计划创建时绑定，而不是按单元格重复解析。
    /// </summary>
    [Fact]
    public void Import_FixedColumnConverter_ShouldBindOnceBeforeCellConversion()
    {
        // Arrange
        var converter = new CountingConverter();
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("订单号");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("a-1");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("b-2");
        }));
        var configuration = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(ImportOrder.OrderNo), ConverterName = "counting" }
            }
        };
        var request = ExcelImport.Workbook<OrderWorkbook>(builder => builder
            .Sheet("Data", root => root.Orders, sheet => sheet.Mapping(configuration)));

        // Act
        var result = new NpoiExcelImporter(valueConverters: new IExcelValueConverter[] { converter })
            .Import(source, request);

        // Assert
        Assert.True(result.IsSuccess, string.Join("; ", result.Errors.Select(error => error.Message)));
        Assert.Equal(2, result.Workbook.Orders.Count);
        Assert.Equal(1, converter.CanConvertCalls);
        Assert.Equal(2, converter.ConvertFromCalls);
    }

    /// <summary>
    /// 测试 - 动态列转换器应在列计划创建时绑定一次，并复用到每个数据行。
    /// </summary>
    [Fact]
    public void Import_DynamicColumnConverter_ShouldBindOnceBeforeCellConversion()
    {
        // Arrange
        var converter = new CountingConverter();
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("地区");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("east");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("west");
        }));
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "地区",
            DataType = typeof(string),
            ConverterName = "counting"
        };
        var request = ExcelImport.Workbook<DynamicTargetWorkbook>(builder =>
            builder.Sheet("Data", root => root.Rows, sheet => sheet
                .DynamicColumns(row => row.TargetFields, new[] { definition })));

        // Act
        var result = new NpoiExcelImporter(valueConverters: new IExcelValueConverter[] { converter })
            .Import(source, request);

        // Assert
        Assert.True(result.IsSuccess, string.Join("; ", result.Errors.Select(error => error.Message)));
        Assert.Equal(2, result.Workbook.Rows.Count);
        Assert.Equal(1, converter.CanConvertCalls);
        Assert.Equal(2, converter.ConvertFromCalls);
    }

    /// <summary>
    /// 测试 - 导出固定列转换器应在列计划创建时绑定并复用到每个数据行。
    /// </summary>
    [Fact]
    public void Export_FixedColumnConverter_ShouldBindOnceBeforeCellConversion()
    {
        // Arrange
        var converter = new CountingConverter();
        var configuration = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(ExportCustomer.Name), ConverterName = "counting" }
            }
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", new[]
        {
            new ExportCustomer { Name = "a" },
            new ExportCustomer { Name = "b" }
        }, sheet => sheet.Mapping(configuration)));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter(new IExcelValueConverter[] { converter }).Export(request, destination);

        // Assert
        Assert.Equal(1, converter.CanConvertCalls);
        Assert.Equal(2, converter.ConvertToCalls);
    }

    /// <summary>
    /// 测试 - Fixed 列宽应按字符宽度写入工作表。
    /// </summary>
    [Fact]
    public void Export_ColumnWidthFixed_ShouldWriteConfiguredWidth()
    {
        // Arrange
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("客户",
            new[] { new ExportCustomer { Name = "客户 A" } }, sheet => sheet
                .ColumnWidth(new ExcelColumnWidthOptions
                {
                    Mode = ExcelColumnWidthMode.Fixed,
                    FixedWidth = 24
                })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        Assert.Equal(24 * 256, result.GetSheet("客户").GetColumnWidth(0));
    }

    /// <summary>
    /// 测试 - 自定义表头批注应写入传统 Note，并保留作者和可见性。
    /// </summary>
    [Fact]
    public void Export_HeaderComment_ShouldWriteLegacyNote()
    {
        // Arrange
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("客户",
            new[] { new ExportCustomer { Name = "客户 A" } }, sheet => sheet
                .HeaderRowIndex(1)
                .DataRowStartIndex(2)
                .HeaderRows(new[] { new ExcelHeaderRow(0,
                    new[] { new ExcelHeaderCell(1, "说明", comment: new ExcelComment("表头说明", "tester", true)) }) })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        var comment = result.GetSheet("客户").GetRow(0).GetCell(1).CellComment;
        Assert.NotNull(comment);
        Assert.Equal("表头说明", comment.String.String);
        Assert.Equal("tester", comment.Author);
        Assert.True(comment.Visible);
    }

    /// <summary>
    /// 测试 - XLSX 应按列 Key 创建柱状、折线和饼图，并拒绝饼图多系列配置。
    /// </summary>
    [Fact]
    public void Export_XlsxCharts_ShouldCreateSupportedChartTypes()
    {
        // Arrange
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("统计",
            new[]
            {
                new ChartRow { Month = "一月", Amount = 10, Forecast = 12 },
                new ChartRow { Month = "二月", Amount = 15, Forecast = 16 }
            }, sheet => sheet
                .Chart(new ExcelChartDefinition
                {
                    Title = "金额柱状图",
                    Type = ExcelChartType.Column,
                    Categories = new ExcelChartRange { ColumnKey = "Month" },
                    Series = new[] { new ExcelChartSeries
                    {
                        Name = "金额", Values = new ExcelChartRange { ColumnKey = "Amount" }
                    } },
                    Anchor = new ExcelChartAnchor { StartRow = 0, StartColumn = 5, EndRow = 12, EndColumn = 14 }
                })
                .Chart(new ExcelChartDefinition
                {
                    Title = "预测折线图",
                    Type = ExcelChartType.Line,
                    Categories = new ExcelChartRange { ColumnKey = "Month" },
                    Series = new[] { new ExcelChartSeries
                    {
                        Name = "预测", Values = new ExcelChartRange { ColumnKey = "Forecast" }
                    } },
                    Anchor = new ExcelChartAnchor { StartRow = 13, StartColumn = 5, EndRow = 25, EndColumn = 14 }
                })
                .Chart(new ExcelChartDefinition
                {
                    Title = "金额饼图",
                    Type = ExcelChartType.Pie,
                    Categories = new ExcelChartRange { ColumnKey = "Month" },
                    Series = new[] { new ExcelChartSeries
                    {
                        Name = "金额", Values = new ExcelChartRange { ColumnKey = "Amount" }
                    } },
                    Anchor = new ExcelChartAnchor { StartRow = 26, StartColumn = 5, EndRow = 38, EndColumn = 14 }
                })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var workbook = (XSSFWorkbook)WorkbookFactory.Create(destination);
        var drawing = Assert.IsType<XSSFDrawing>(workbook.GetSheet("统计").CreateDrawingPatriarch());
        Assert.Equal(3, drawing.GetCharts().Count);
        Assert.Contains(drawing.GetCharts(), chart => chart.Title.String == "金额柱状图");
        Assert.Contains(drawing.GetCharts(), chart => chart.Title.String == "预测折线图");
        Assert.Contains(drawing.GetCharts(), chart => chart.Title.String == "金额饼图");
    }

    private static byte[] CreateWorkbook(Action<IWorkbook> configure)
    {
        using var workbook = new XSSFWorkbook();
        configure(workbook);
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    private sealed class ExportOrder
    {
        public string OrderNo { get; set; }

        [DynamicColumn]
        public IDictionary<string, object> CustomFields { get; set; }
    }

    private sealed class ExportCustomer
    {
        public string Name { get; set; }
    }

    private sealed class ChartRow
    {
        public string Month { get; set; }
        public double Amount { get; set; }
        public double Forecast { get; set; }
    }

    private sealed class OrderWorkbook
    {
        public List<ImportOrder> Orders { get; } = new List<ImportOrder>();
        public List<ImportDetail> Details { get; } = new List<ImportDetail>();
    }

    private sealed class ImportOrder
    {
        [ColumnName("订单号")]
        public string OrderNo { get; set; }

        [DynamicColumn]
        public IDictionary<string, object> CustomFields { get; set; }

        [ExcelIgnore]
        public List<ImportDetail> DetailItems { get; } = new List<ImportDetail>();
    }

    private sealed class ImportDetail
    {
        [ColumnName("订单号")]
        public string OrderNo { get; set; }
        [ColumnName("名称")]
        public string Name { get; set; }
    }

    private sealed class IntRelationWorkbook
    {
        public List<IntParent> Parents { get; } = new();
        public List<IntChild> Children { get; } = new();
    }

    private sealed class IntParent
    {
        public int Id { get; set; }

        [ExcelIgnore]
        public List<IntChild> Children { get; } = new();
    }

    private sealed class IntChild
    {
        public int ParentId { get; set; }
    }

    private sealed class SelectorWorkbook
    {
        public List<SelectorRow> Rows { get; } = new List<SelectorRow>();
    }

    private sealed class SelectorRow
    {
        public string Code { get; set; }
    }

    private sealed class DynamicTargetWorkbook
    {
        public List<DynamicTargetRow> Rows { get; } = new List<DynamicTargetRow>();
    }

    private sealed class DynamicTargetRow
    {
        [DynamicColumn]
        public IDictionary<string, object> MappedFields { get; set; }

        [ExcelIgnore]
        public IDictionary<string, object> TargetFields { get; } =
            new Dictionary<string, object>(StringComparer.Ordinal);
    }

    private sealed class UpperRegionConverter : INamedExcelValueConverter
    {
        public string Name => "upper-region";

        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = Convert.ToString(context.Value)?.ToUpperInvariant();
            return true;
        }

        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = Convert.ToString(context.Value)?.ToUpperInvariant();
            return true;
        }
    }

    private sealed class CountingConverter : INamedExcelValueConverter
    {
        public string Name => "counting";

        public int CanConvertCalls { get; private set; }

        public int ConvertFromCalls { get; private set; }

        public int ConvertToCalls { get; private set; }

        public bool CanConvert(Type propertyType)
        {
            CanConvertCalls++;
            return propertyType == typeof(string);
        }

        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            ConvertFromCalls++;
            value = Convert.ToString(context.Value)?.ToUpperInvariant();
            return true;
        }

        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            ConvertToCalls++;
            value = context.Value;
            return true;
        }
    }
}
