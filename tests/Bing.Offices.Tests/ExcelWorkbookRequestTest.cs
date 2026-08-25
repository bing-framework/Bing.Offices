using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Csv;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Npoi.Exports;
using Bing.Offices.Npoi.Imports;
using Bing.Offices.Npoi;
using Bing.Offices.Providers;
using Bing.Offices.Styles;
using Bing.Offices.Validations;
using System.Text;
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
    /// 测试 - 同一映射的多个 Sheet 应由一次 Workbook Plan 构建覆盖，并实际写入两个 Sheet。
    /// </summary>
    [Fact]
    public void Export_WorkbookPlan_ShouldBuildOnceForSameMappingSheets()
    {
        // Arrange
        var recorder = new RecordingMappingPlanFactory();
        var request = ExcelExport.Workbook(workbook =>
        {
            workbook.AddSheet("订单一", new[] { new ExportOrder { OrderNo = "O-1" } });
            workbook.AddSheet("订单二", new[] { new ExportOrder { OrderNo = "O-2" } });
        });
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter(mappingPlanFactory: recorder).Export(request, destination);

        // Assert
        Assert.Equal(new[] { 2 }, recorder.WorkbookPlanSheetCounts);
        destination.Position = 0;
        using var workbookResult = WorkbookFactory.Create(destination);
        Assert.Equal("O-1", workbookResult.GetSheet("订单一").GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("O-2", workbookResult.GetSheet("订单二").GetRow(1).GetCell(0).StringCellValue);
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
        var map = new Bing.Offices.Mappings.ExcelMappingPlanFactory().Create<ImportOrder>(new ExcelMappingDocument(), null,
            MappingDirection.Import);
        var fixedProperty = map.Columns.Single(property => property.Name == nameof(ImportOrder.OrderNo));
        var dynamicProperty = map.Columns.Single(property => property.Name == nameof(ImportOrder.CustomFields));
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "地区",
            DataType = typeof(string),
            ConverterName = "region",
            ValidatorName = "required"
        };
        var fixedPlan = new ExcelColumnPlan(fixedProperty.Title, fixedProperty, false, 0, null,
            reflectionProperty: typeof(ImportOrder).GetProperty(nameof(ImportOrder.OrderNo)));
        var dynamicPlan = new ExcelColumnPlan(definition.Title, dynamicProperty, true, 1, definition,
            definition.Key, reflectionProperty: typeof(ImportOrder).GetProperty(nameof(ImportOrder.CustomFields)));

        // Assert
        Assert.NotNull(fixedPlan.Getter);
        Assert.NotNull(fixedPlan.Setter);
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

        var legacyDocument = new ExcelMappingDocument
        {
            Export = new ExcelMappingConfiguration
            {
                DynamicColumns =
                {
                    new ExcelMappingDynamicColumnConfiguration
                    {
                        Key = "region", Title = "地区", DataTypeName = "string",
                        PlacementKey = "before-OrderNo"
                    }
                }
            }
        };
        using var legacyDestination = new MemoryStream();
        new NpoiExcelExporter().Export(ExcelExport.Workbook(workbook => workbook.AddSheet("订单",
            new[] { new ExportOrder { OrderNo = "O-1", CustomFields = new Dictionary<string, object>
            {
                ["region"] = "华东"
            } } }, sheet => sheet.Mapping(legacyDocument))), legacyDestination);
        legacyDestination.Position = 0;
        using var legacyWorkbook = WorkbookFactory.Create(legacyDestination);
        Assert.Equal("地区", legacyWorkbook.GetSheet("订单").GetRow(0).GetCell(0).StringCellValue);
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
    /// 测试 - Document 动态列、布局、数字格式和样式应进入真实 XLSX 导出计划。
    /// </summary>
    [Fact]
    public void Export_DocumentDynamicPlan_ShouldApplyLayoutFormatAndStyle()
    {
        // Arrange
        var document = new ExcelMappingDocument
        {
            Export = new ExcelMappingConfiguration
            {
                DynamicColumns =
                {
                    new ExcelMappingDynamicColumnConfiguration
                    {
                        Key = "amount", Title = "金额", DataTypeName = "decimal",
                        NumberFormat = "0.00", PlacementKey = "before:OrderNo"
                    }
                },
                Style = new ExcelMappingStyleConfiguration { HeaderStyleKey = "header" }
            }
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("订单", new[]
        {
            new ExportOrder
            {
                OrderNo = "O-1",
                CustomFields = new Dictionary<string, object> { ["amount"] = 12.5m }
            }
        }, sheet => sheet.Mapping(document)));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var sheet = workbook.GetSheet("订单");
        Assert.Equal("金额", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.True(sheet.GetRow(0).GetCell(0).CellStyle.GetFont(workbook).IsBold);
        Assert.Equal(CellType.Numeric, sheet.GetRow(1).GetCell(0).CellType);
        Assert.Equal("0.00", sheet.GetRow(1).GetCell(0).CellStyle.GetDataFormatString());
        Assert.Equal("O-1", sheet.GetRow(1).GetCell(1).StringCellValue);
    }

    /// <summary>
    /// 测试 - Document 命名动态校验器应在 CSV 与 XLSX 主链产生等价失败。
    /// </summary>
    [Fact]
    public void Import_DocumentDynamicValidator_ShouldMatchCsvAndXlsx()
    {
        // Arrange
        var document = new ExcelMappingDocument
        {
            Import = new ExcelMappingConfiguration
            {
                DynamicColumns =
                {
                    new ExcelMappingDynamicColumnConfiguration
                    {
                        Key = "region", Title = "地区", DataTypeName = "string",
                        ValidatorName = "starts-ok",
                        ValidationRuleNames = new List<string> { "contains-region" }
                    }
                }
            }
        };
        var validator = new StartsWithOkValidationRule();
        var secondValidator = new ContainsRegionValidationRule();
        using var xlsx = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("地区");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("BAD");
        }));
        var request = ExcelImport.Workbook<DynamicTargetWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet => sheet.Mapping(document)));
        using var csv = new MemoryStream(Encoding.UTF8.GetBytes("地区\r\nBAD\r\n"));

        // Act
        var validators = new INamedExcelValidationRule[] { validator, secondValidator };
        var plan = new ExcelMappingPlanFactory(namedValidationRules: validators)
            .Create<DynamicTargetRow>(document, MappingDirection.Import);
        var dynamicPlan = Assert.Single(plan.DynamicColumns);
        Assert.Equal(new[] { "starts-ok", "contains-region" }, dynamicPlan.ValidationRuleNames);
        Assert.Equal(2, dynamicPlan.ValidationBindings.Count);
        var xlsxResult = new NpoiExcelImporter(namedValidationRules: validators).Import(xlsx, request);
        var csvResult = new CsvEntityImporter(namedValidationRules: validators).Import<DynamicTargetRow>(csv,
            new CsvImportOptions<DynamicTargetRow> { MappingDocument = document, HeaderMatch = false });

        // Assert
        Assert.Single(xlsxResult.Errors);
        Assert.Single(csvResult.Errors);
        Assert.Equal(2, xlsxResult.Errors[0].RowIndex);
        Assert.Equal(2, csvResult.Errors[0].RowIndex);
        Assert.Empty(xlsxResult.Workbook.Rows);
        Assert.Empty(csvResult.Items);
    }

    /// <summary>
    /// 测试 - 七类内置动态规则应在 CSV 与 XLSX 中产生一致失败行和唯一值回滚结果。
    /// </summary>
    [Fact]
    public void Import_DynamicBuiltInValidationMatrix_ShouldMatchCsvAndXlsx()
    {
        // Arrange
        var document = new ExcelMappingDocument
        {
            Import = new ExcelMappingConfiguration
            {
                DynamicColumns = CreateDynamicValidationColumns()
            }
        };
        var headers = new[] { "必填", "正则", "日期", "最大值", "区间", "长度", "唯一" };
        var rows = new[]
        {
            new[] { "ok", "OK-1", "2026-08-24", "5", "5", "abc", "U-1" },
            new[] { "ok", "OK-1", "2026-08-24", "10", "1", "abcde", "U-2" },
            new[] { "ok", "OK-1", "2026-08-24", "10", "10", "abcde", "U-3" },
            new[] { "", "OK-1", "2026-08-24", "5", "5", "abc", "U-4" },
            new[] { "ok", "BAD", "2026-08-24", "5", "5", "abc", "U-5" },
            new[] { "ok", "OK-1", "not-date", "5", "5", "abc", "U-6" },
            new[] { "ok", "OK-1", "2026-08-24", "11", "5", "abc", "U-7" },
            new[] { "ok", "OK-1", "2026-08-24", "5", "0", "abc", "U-8" },
            new[] { "ok", "OK-1", "2026-08-24", "5", "5", "abcdef", "U-9" },
            new[] { "ok", "OK-1", "2026-08-24", "5", "5", "abc", "u-1" }
        };
        using var xlsx = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0);
            for (var column = 0; column < headers.Length; column++)
                sheet.GetRow(0).CreateCell(column).SetCellValue(headers[column]);
            for (var row = 0; row < rows.Length; row++)
            {
                sheet.CreateRow(row + 1);
                for (var column = 0; column < rows[row].Length; column++)
                    sheet.GetRow(row + 1).CreateCell(column).SetCellValue(rows[row][column]);
            }
        }));
        var request = ExcelImport.Workbook<DynamicTargetWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet => sheet.Mapping(document)));
        using var csv = new MemoryStream(Encoding.UTF8.GetBytes(
            string.Join("\r\n", new[] { string.Join(",", headers) }
                .Concat(rows.Select(row => string.Join(",", row)))) + "\r\n"));

        // Act
        var plan = new ExcelMappingPlanFactory().Create<DynamicTargetRow>(document, MappingDirection.Import);
        var xlsxResult = new NpoiExcelImporter().Import(xlsx, request);
        var csvResult = new CsvEntityImporter().Import<DynamicTargetRow>(csv,
            new CsvImportOptions<DynamicTargetRow> { MappingDocument = document, HeaderMatch = false });

        // Assert
        Assert.Equal(7, plan.DynamicColumns.Count);
        Assert.All(plan.DynamicColumns, column => Assert.Single(column.ValidationBindings));
        Assert.True(plan.DynamicColumns.Single(column => column.Key == "unique").IsUnique);
        Assert.Equal(7, xlsxResult.Errors.Count);
        Assert.Equal(7, csvResult.Errors.Count);
        Assert.Equal(3, xlsxResult.Workbook.Rows.Count);
        Assert.Equal(3, csvResult.Items.Count);
        var expected = new[]
        {
            (Row: 5, Column: 1, Key: "required", Code: ExcelImportErrorCode.Validation),
            (Row: 6, Column: 2, Key: "regex", Code: ExcelImportErrorCode.Validation),
            (Row: 7, Column: 3, Key: "date", Code: ExcelImportErrorCode.ValueConversion),
            (Row: 8, Column: 4, Key: "max", Code: ExcelImportErrorCode.MaxValue),
            (Row: 9, Column: 5, Key: "range", Code: ExcelImportErrorCode.Validation),
            (Row: 10, Column: 6, Key: "length", Code: ExcelImportErrorCode.MaxLength),
            (Row: 11, Column: 7, Key: "unique", Code: ExcelImportErrorCode.Validation)
        };
        Assert.Equal(expected.Select(item => $"{item.Code}|{item.Row}|{item.Column}|{item.Key}"),
            xlsxResult.Errors.Select(error => $"{error.Code}|{error.RowIndex}|{error.ColumnIndex}|{error.ColumnKey}"));
        Assert.Equal(expected.Select(item => $"{item.Row}|{item.Column}|{item.Key}"),
            csvResult.Errors.Select(error => $"{error.RowIndex}|{error.ColumnIndex}|{error.PropertyName}"));
        Assert.Equal(2, xlsxResult.Errors.Single(error => error.ColumnKey == "unique").FirstRowNumber);
        Assert.Equal(2, csvResult.Errors.Single(error => error.PropertyName == "unique").FirstRowNumber);
    }

    /// <summary>
    /// 测试 - 动态唯一规则应在 CSV 与 XLSX 中一致处理空值、大小写比较和跟踪上限。
    /// </summary>
    [Fact]
    public void Import_DynamicUniqueOptions_ShouldMatchCsvAndXlsx()
    {
        // Arrange
        var values = new[] { " ", " ", "ABC", "abc" };
        byte[] CreateSource() => CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("唯一");
            sheet.GetRow(0).CreateCell(1).SetCellValue("辅助");
            for (var row = 0; row < values.Length; row++)
            {
                sheet.CreateRow(row + 1).CreateCell(0).SetCellValue(values[row]);
                sheet.GetRow(row + 1).CreateCell(1).SetCellValue("x");
            }
        });
        var csvText = "唯一,辅助\r\n,x\r\n,x\r\nABC,x\r\nabc,x\r\n";

        // Act
        var ignoredDocument = CreateDynamicUniqueDocument(ignoreEmpty: true);
        var ignoredXlsx = ImportDynamicUniqueXlsx(CreateSource(), ignoredDocument,
            new ExcelResourceLimits { UniqueComparison = StringComparison.OrdinalIgnoreCase });
        var ignoredCsv = ImportDynamicUniqueCsv(csvText, ignoredDocument,
            new CsvImportOptions<DynamicTargetRow>
            {
                HeaderMatch = false,
                UniqueComparison = StringComparison.OrdinalIgnoreCase
            });

        var trackedDocument = CreateDynamicUniqueDocument(ignoreEmpty: false);
        var trackedXlsx = ImportDynamicUniqueXlsx(CreateSource(), trackedDocument,
            new ExcelResourceLimits { UniqueComparison = StringComparison.OrdinalIgnoreCase });
        var trackedCsv = ImportDynamicUniqueCsv(csvText, trackedDocument,
            new CsvImportOptions<DynamicTargetRow>
            {
                HeaderMatch = false,
                UniqueComparison = StringComparison.OrdinalIgnoreCase
            });

        var limitedValues = new[] { "ABC", "DEF" };
        var limitedSource = CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("唯一");
            for (var row = 0; row < limitedValues.Length; row++)
                sheet.CreateRow(row + 1).CreateCell(0).SetCellValue(limitedValues[row]);
        });
        var limitedDocument = CreateDynamicUniqueDocument(ignoreEmpty: true);
        var limitedXlsx = ImportDynamicUniqueXlsx(limitedSource, limitedDocument,
            new ExcelResourceLimits
            {
                MaxTrackedUniqueValues = 1,
                UniqueComparison = StringComparison.OrdinalIgnoreCase
            });
        var limitedCsv = ImportDynamicUniqueCsv("唯一\r\nABC\r\nDEF\r\n", limitedDocument,
            new CsvImportOptions<DynamicTargetRow>
            {
                HeaderMatch = false,
                MaxTrackedUniqueValues = 1,
                UniqueComparison = StringComparison.OrdinalIgnoreCase
            });

        // Assert
        Assert.Equal(1, ignoredXlsx.Errors.Count);
        Assert.Equal(1, ignoredCsv.Errors.Count);
        Assert.Equal(3, ignoredXlsx.Workbook.Rows.Count);
        Assert.Equal(3, ignoredCsv.Items.Count);
        Assert.Equal(5, ignoredXlsx.Errors[0].RowIndex);
        Assert.Equal(5, ignoredCsv.Errors[0].RowIndex);
        Assert.Equal(4, ignoredXlsx.Errors[0].FirstRowNumber);
        Assert.Equal(4, ignoredCsv.Errors[0].FirstRowNumber);

        Assert.Equal(2, trackedXlsx.Errors.Count);
        Assert.Equal(2, trackedCsv.Errors.Count);
        Assert.Equal(2, trackedXlsx.Workbook.Rows.Count);
        Assert.Equal(2, trackedCsv.Items.Count);
        Assert.Equal(new[] { 3, 5 }, trackedXlsx.Errors.Select(error => error.RowIndex));
        Assert.Equal(new[] { 3, 5 }, trackedCsv.Errors.Select(error => error.RowIndex));
        Assert.Equal(new int?[] { 2, 4 }, trackedXlsx.Errors.Select(error => error.FirstRowNumber));
        Assert.Equal(new int?[] { 2, 4 }, trackedCsv.Errors.Select(error => error.FirstRowNumber));
        Assert.All(trackedXlsx.Errors, error => Assert.Equal("unique", error.ColumnKey));
        Assert.All(trackedCsv.Errors, error => Assert.Equal("unique", error.PropertyName));

        Assert.Equal(1, limitedXlsx.Errors.Count);
        Assert.Equal(1, limitedCsv.Errors.Count);
        Assert.Equal(3, limitedXlsx.Errors[0].RowIndex);
        Assert.Equal(3, limitedCsv.Errors[0].RowIndex);
        Assert.Equal("unique", limitedXlsx.Errors[0].ColumnKey);
        Assert.Equal("unique", limitedCsv.Errors[0].PropertyName);
    }

    /// <summary>
    /// 测试 - 七类固定属性规则应在 CSV 与 XLSX 中保持相同 binding、坐标和失败行回滚契约。
    /// </summary>
    [Fact]
    public void Import_FixedBuiltInValidationMatrix_ShouldMatchCsvAndXlsx()
    {
        // Arrange
        var document = new ExcelMappingDocument
        {
            Import = new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    new() { PropertyName = nameof(FixedValidationRow.Required), Title = "必填" },
                    new() { PropertyName = nameof(FixedValidationRow.Regex), Title = "正则" },
                    new() { PropertyName = nameof(FixedValidationRow.Date), Title = "日期" },
                    new() { PropertyName = nameof(FixedValidationRow.MaxValue), Title = "最大值" },
                    new() { PropertyName = nameof(FixedValidationRow.Range), Title = "区间" },
                    new() { PropertyName = nameof(FixedValidationRow.MaxLength), Title = "长度" },
                    new() { PropertyName = nameof(FixedValidationRow.Unique), Title = "唯一" }
                }
            }
        };
        var headers = new[] { "必填", "正则", "日期", "最大值", "区间", "长度", "唯一" };
        var rows = new[]
        {
            new[] { "ok", "OK-1", "2026-08-24", "5", "5", "abc", "U-1" },
            new[] { "ok", "OK-1", "2026-08-24", "10", "1", "abcde", "U-2" },
            new[] { "ok", "OK-1", "2026-08-24", "10", "10", "abcde", "U-3" },
            new[] { "", "OK-1", "2026-08-24", "5", "5", "abc", "U-4" },
            new[] { "ok", "BAD", "2026-08-24", "5", "5", "abc", "U-5" },
            new[] { "ok", "OK-1", "not-date", "5", "5", "abc", "U-6" },
            new[] { "ok", "OK-1", "2026-08-24", "11", "5", "abc", "U-7" },
            new[] { "ok", "OK-1", "2026-08-24", "5", "11", "abc", "U-8" },
            new[] { "ok", "OK-1", "2026-08-24", "5", "5", "abcdef", "U-9" },
            new[] { "ok", "OK-1", "2026-08-24", "5", "5", "abc", "u-1" }
        };
        using var xlsx = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0);
            for (var column = 0; column < headers.Length; column++)
                sheet.GetRow(0).CreateCell(column).SetCellValue(headers[column]);
            for (var row = 0; row < rows.Length; row++)
            {
                sheet.CreateRow(row + 1);
                for (var column = 0; column < rows[row].Length; column++)
                    sheet.GetRow(row + 1).CreateCell(column).SetCellValue(rows[row][column]);
            }
        }));
        var request = ExcelImport.Workbook<FixedValidationWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet => sheet.Mapping(document)));
        using var csv = new MemoryStream(Encoding.UTF8.GetBytes(
            string.Join("\r\n", new[] { string.Join(",", headers) }
                .Concat(rows.Select(row => string.Join(",", row)))) + "\r\n"));

        // Act
        var xlsxResult = new NpoiExcelImporter().Import(xlsx, request);
        var csvResult = new CsvEntityImporter().Import<FixedValidationRow>(csv,
            new CsvImportOptions<FixedValidationRow> { MappingDocument = document, HeaderMatch = false });

        // Assert
        Assert.Equal(7, xlsxResult.Errors.Count);
        Assert.Equal(7, csvResult.Errors.Count);
        Assert.Equal(3, xlsxResult.Workbook.Rows.Count);
        Assert.Equal(3, csvResult.Items.Count);
        var expected = new[]
        {
            (Row: 5, Column: 1, Property: nameof(FixedValidationRow.Required), Code: ExcelImportErrorCode.Validation),
            (Row: 6, Column: 2, Property: nameof(FixedValidationRow.Regex), Code: ExcelImportErrorCode.Validation),
            (Row: 7, Column: 3, Property: nameof(FixedValidationRow.Date), Code: ExcelImportErrorCode.ValueConversion),
            (Row: 8, Column: 4, Property: nameof(FixedValidationRow.MaxValue), Code: ExcelImportErrorCode.MaxValue),
            (Row: 9, Column: 5, Property: nameof(FixedValidationRow.Range), Code: ExcelImportErrorCode.Validation),
            (Row: 10, Column: 6, Property: nameof(FixedValidationRow.MaxLength), Code: ExcelImportErrorCode.MaxLength),
            (Row: 11, Column: 7, Property: nameof(FixedValidationRow.Unique), Code: ExcelImportErrorCode.Validation)
        };
        Assert.Equal(expected.Select(item => $"{item.Code}|{item.Row}|{item.Column}|{item.Property}"),
            xlsxResult.Errors.Select(error => $"{error.Code}|{error.RowIndex}|{error.ColumnIndex}|{error.PropertyName}"));
        Assert.Equal(expected.Select(item => $"{item.Row}|{item.Column}|{item.Property}"),
            csvResult.Errors.Select(error => $"{error.RowIndex}|{error.ColumnIndex}|{error.PropertyName}"));
        Assert.Equal(2, xlsxResult.Errors.Single(error => error.PropertyName == nameof(FixedValidationRow.Unique)).FirstRowNumber);
        Assert.Equal(2, csvResult.Errors.Single(error => error.PropertyName == nameof(FixedValidationRow.Unique)).FirstRowNumber);
    }

    private static List<ExcelMappingDynamicColumnConfiguration> CreateDynamicValidationColumns() => new()
    {
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "required", Title = "必填", DataTypeName = "string", Order = 1,
            ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
            {
                new() { Name = "required" }
            }
        },
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "regex", Title = "正则", DataTypeName = "string", Order = 2,
            ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
            {
                new() { Name = "regex", Pattern = "^OK-" }
            }
        },
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "date", Title = "日期", DataTypeName = "datetime", Order = 3,
            ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
            {
                new() { Name = "date", Format = "yyyy-MM-dd" }
            }
        },
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "max", Title = "最大值", DataTypeName = "decimal", Order = 4,
            ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
            {
                new() { Name = "maxValue", MaxValue = 10 }
            }
        },
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "range", Title = "区间", DataTypeName = "decimal", Order = 5,
            ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
            {
                new() { Name = "range", Min = 1, Max = 10 }
            }
        },
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "length", Title = "长度", DataTypeName = "string", Order = 6,
            ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
            {
                new() { Name = "maxLength", MaxLength = 5 }
            }
        },
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "unique", Title = "唯一", DataTypeName = "string", Order = 7,
            ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
            {
                new() { Name = "unique", IgnoreEmpty = false }
            }
        }
    };

    private static ExcelMappingDocument CreateDynamicUniqueDocument(bool ignoreEmpty) => new()
    {
        Import = new ExcelMappingConfiguration
        {
            DynamicColumns =
            {
                new ExcelMappingDynamicColumnConfiguration
                {
                    Key = "unique", Title = "唯一", DataTypeName = "string",
                    ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
                    {
                        new() { Name = "unique", IgnoreEmpty = ignoreEmpty }
                    }
                }
            }
        }
    };

    private static ExcelWorkbookImportResult<DynamicTargetWorkbook> ImportDynamicUniqueXlsx(
        byte[] source, ExcelMappingDocument document, ExcelResourceLimits resourceLimits)
    {
        using var stream = new MemoryStream(source);
        var request = ExcelImport.Workbook<DynamicTargetWorkbook>(workbook => workbook
            .ResourceLimits(resourceLimits)
            .Sheet("Data", root => root.Rows, sheet => sheet.Mapping(document)));
        return new NpoiExcelImporter().Import(stream, request);
    }

    private static CsvImportResult<DynamicTargetRow> ImportDynamicUniqueCsv(string source,
        ExcelMappingDocument document, CsvImportOptions<DynamicTargetRow> options)
    {
        options.MappingDocument = document;
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(source));
        return new CsvEntityImporter().Import(stream, options);
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

    private sealed class FixedValidationWorkbook
    {
        public List<FixedValidationRow> Rows { get; } = new();
    }

    private sealed class FixedValidationRow
    {
        [ExcelRequired]
        public string Required { get; set; }

        [ExcelRegex("^OK-")]
        public string Regex { get; set; }

        [ExcelDate("yyyy-MM-dd")]
        public DateTime Date { get; set; }

        [ExcelMaxValue(10)]
        public decimal MaxValue { get; set; }

        [ExcelRange(1, 10)]
        public decimal Range { get; set; }

        [ExcelMaxLength(5)]
        public string MaxLength { get; set; }

        [ExcelUnique(IgnoreEmpty = false)]
        public string Unique { get; set; }
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

    private sealed class StartsWithOkValidationRule : INamedExcelValidationRule
    {
        public string Name => "starts-ok";
        public string ErrorMessage => "必须以 OK 开头";
        public bool Validate(ExcelValidationContext context) => context.Value.StartsWith("OK",
            StringComparison.Ordinal);
    }

    private sealed class ContainsRegionValidationRule : INamedExcelValidationRule
    {
        public string Name => "contains-region";
        public string ErrorMessage => "必须包含地区";
        public bool Validate(ExcelValidationContext context) => context.Value.Contains("地区",
            StringComparison.Ordinal);
    }

    private sealed class RecordingMappingPlanFactory : IExcelMappingPlanFactory
    {
        private readonly ExcelMappingPlanFactory _inner = new();
        private readonly List<int> _workbookPlanSheetCounts = new();

        public IReadOnlyList<int> WorkbookPlanSheetCounts => _workbookPlanSheetCounts;

        public IExcelMappingPlan Create<T>(ExcelMappingDocument document, MappingDirection direction)
            where T : class, new() => _inner.Create<T>(document, direction);

        public IExcelMappingWorkbookPlan CreateWorkbook<T>(ExcelMappingDocument document,
            MappingDirection direction, IReadOnlyList<string> sheetNames) where T : class, new()
        {
            _workbookPlanSheetCounts.Add(sheetNames.Count);
            return _inner.CreateWorkbook<T>(document, direction, sheetNames);
        }
    }
}
