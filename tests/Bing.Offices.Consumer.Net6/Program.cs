using System.Text;
using Bing.Offices;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.Csv;
using Bing.Offices.Entities;
using Bing.Offices.ExcelDataReader.Extensions;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.MiniExcel.Extensions;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using NPOI.XSSF.UserModel;

var directory = Path.Combine(Path.GetTempPath(), "Bing.Offices.Consumer", Guid.NewGuid().ToString("N"));
var packageVersion = Environment.GetEnvironmentVariable("BING_OFFICES_PACKAGE_VERSION");
if (string.IsNullOrWhiteSpace(packageVersion))
    throw new InvalidOperationException("BING_OFFICES_PACKAGE_VERSION must identify the packages under test.");
Directory.CreateDirectory(directory);

try
{
    var services = new ServiceCollection();
    services.AddBingOfficesNpoi();
    using var provider = services.BuildServiceProvider();
    var csvExporter = provider.GetRequiredService<ICsvExporter>();
    var csvImporter = provider.GetRequiredService<ICsvImporter>();
    var excelExporter = provider.GetRequiredService<IExcelExporter>();
    var excelImporter = provider.GetRequiredService<IExcelImporter>();
    await Bing.Offices.PackageConsumer.BasicPivotPackageContract.VerifyAsync(excelExporter);
    await Bing.Offices.PackageConsumer.AdvancedPivotPackageContract.VerifyAsync(excelExporter);
    await Bing.Offices.PackageConsumer.AdvancedPivotPackageContract.VerifyTemplateUpdateAsync(excelExporter);
    var rows = new[] { new ConsumerRow { Code = "consumer,net6", Count = 7 } };
    var csvOptions = new CsvExportOptions<ConsumerRow> { Encoding = new UTF8Encoding(false, true) };

    var csvSyncBytes = CsvStreamExtensions.ExportToBytes(csvExporter, rows, csvOptions);
    var csvSyncResult = CsvStreamExtensions.ImportFromBytes<ConsumerRow>(csvImporter, csvSyncBytes,
        new CsvImportOptions<ConsumerRow> { Encoding = new UTF8Encoding(false, true) });
    var csvAsyncBytes = await CsvStreamExtensions.ExportToBytesAsync(csvExporter, rows, csvOptions);
    var csvAsyncResult = await CsvStreamExtensions.ImportFromBytesAsync<ConsumerRow>(csvImporter,
        csvAsyncBytes, new CsvImportOptions<ConsumerRow> { Encoding = new UTF8Encoding(false, true) });

    var csvSyncPath = Path.Combine(directory, "sync.csv");
    csvExporter.ExportToFile(rows, csvSyncPath, csvOptions);
    var csvFileResult = CsvStreamExtensions.ImportFromFile<ConsumerRow>(csvImporter, csvSyncPath,
        new CsvImportOptions<ConsumerRow> { Encoding = new UTF8Encoding(false, true) });
    var csvAsyncPath = Path.Combine(directory, "async.csv");
    await csvExporter.ExportToFileAsync(rows, csvAsyncPath, csvOptions);
    var csvAsyncFileResult = await CsvStreamExtensions.ImportFromFileAsync<ConsumerRow>(csvImporter,
        csvAsyncPath, new CsvImportOptions<ConsumerRow> { Encoding = new UTF8Encoding(false, true) });

    var workbookRequest = ExcelExport.Workbook(workbook => workbook
        .Format(ExcelFormat.Xlsx)
        .AddSheet("Data", rows));
    var importRequest = ExcelImport.Workbook<ConsumerWorkbook>(workbook =>
        workbook.Sheet("Data", root => root.Rows));
    var excelSyncBytes = ExcelStreamExtensions.ExportToBytes(excelExporter, workbookRequest);
    var excelSyncResult = ExcelStreamExtensions.ImportFromBytes(excelImporter, excelSyncBytes, importRequest);
    var excelAsyncBytes = await ExcelStreamExtensions.ExportToBytesAsync(excelExporter, workbookRequest);
    var excelAsyncResult = await ExcelStreamExtensions.ImportFromBytesAsync(excelImporter,
        excelAsyncBytes, importRequest);

    var excelSyncPath = Path.Combine(directory, "sync.xlsx");
    excelExporter.ExportToFile(workbookRequest, excelSyncPath);
    var excelFileResult = ExcelStreamExtensions.ImportFromFile<ConsumerWorkbook>(excelImporter,
        excelSyncPath, importRequest);
    var excelAsyncPath = Path.Combine(directory, "async.xlsx");
    await excelExporter.ExportToFileAsync(workbookRequest, excelAsyncPath);
    var excelAsyncFileResult = await ExcelStreamExtensions.ImportFromFileAsync<ConsumerWorkbook>(
        excelImporter, excelAsyncPath, importRequest);

    var directExporter = new Bing.Offices.Exports.NpoiExcelExporter();
    var directImporter = new Bing.Offices.Imports.NpoiExcelImporter();
    var directBytes = await directExporter.ExportToBytesAsync(workbookRequest);
    var directResult = await directImporter.ImportFromBytesAsync(directBytes, importRequest);

    using var miniProvider = new ServiceCollection()
        .AddBingOfficesMiniExcel()
        .BuildServiceProvider();
    var miniExporter = miniProvider.GetRequiredService<IExcelExporter>();
    var miniImporter = miniProvider.GetRequiredService<IExcelImporter>();
    var miniBytes = await ExcelStreamExtensions.ExportToBytesAsync(miniExporter, workbookRequest);
    var miniResult = await ExcelStreamExtensions.ImportFromBytesAsync(miniImporter, miniBytes, importRequest);

    using var closedXmlProvider = new ServiceCollection()
        .AddBingOfficesClosedXml()
        .BuildServiceProvider();
    var closedXmlExporter = closedXmlProvider.GetRequiredService<IExcelExporter>();
    var closedXmlImporter = closedXmlProvider.GetRequiredService<IExcelImporter>();
    await Bing.Offices.PackageConsumer.AdvancedPivotPackageContract.VerifyAsync(closedXmlExporter);
    await Bing.Offices.PackageConsumer.AdvancedPivotPackageContract.VerifyTemplateUpdateAsync(closedXmlExporter);
    var closedXmlBytes = await ExcelStreamExtensions.ExportToBytesAsync(closedXmlExporter, workbookRequest);
    var closedXmlResult = await ExcelStreamExtensions.ImportFromBytesAsync(
        closedXmlImporter, closedXmlBytes, importRequest);

    var pivotRequest = ExcelExport.Workbook(workbook => workbook
        .LookupList("PivotCodes", rows.Select(row => row.Code).ToArray())
        .AddSheet("PivotData", rows, sheet => sheet.DataValidation(new ExcelDataValidationDefinition
        {
            Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 1, StartColumn = 0, EndColumn = 0 },
            Type = ExcelDataValidationType.List, ListSourceName = "PivotCodes"
        }))
        .PivotTable(new ExcelPivotTableDefinition
        {
            Name = "ConsumerPivot", SourceSheet = "PivotData",
            SourceRange = new ExcelRangeDefinition
            {
                StartRow = 0, StartColumn = 0, EndRow = 1, EndColumn = 1
            },
            TargetSheet = "__BingLookup_001", TargetRow = 0, TargetColumn = 0,
            RowField = "Code", ValueField = "Count"
        }));
    var pivotBytes = await ExcelStreamExtensions.ExportToBytesAsync(closedXmlExporter, pivotRequest);
    using (var pivotWorkbook = new ClosedXML.Excel.XLWorkbook(new MemoryStream(pivotBytes)))
    {
        Ensure(pivotWorkbook.Worksheets.Select(sheet => sheet.Name).SequenceEqual(new[]
            { "PivotData", "__BingLookup_002", "__BingLookup_001" }),
            "Pivot/lookup sheet allocation did not reserve the future target.");
        Ensure(pivotWorkbook.Worksheet("PivotData").Range("A1:B2").Cells()
                .Select(cell => cell.GetString()).SequenceEqual(new[] { "Code", "Count", rows[0].Code, "7" }),
            "Pivot package source data mismatch.");
        var helper = pivotWorkbook.Worksheet("__BingLookup_002");
        Ensure(helper.Visibility == ClosedXML.Excel.XLWorksheetVisibility.Hidden
            && helper.Cell("A1").GetString() == rows[0].Code && !helper.Cell("A1").HasFormula,
            "Pivot lookup text or hidden state mismatch.");
        var namedRange = pivotWorkbook.DefinedNames.Single(name => name.Name == "PivotCodes").Ranges.Single();
        Ensure(namedRange.Worksheet.Name == helper.Name
            && namedRange.RangeAddress.FirstAddress.RowNumber == 1 && namedRange.RangeAddress.FirstAddress.ColumnNumber == 1
            && namedRange.RangeAddress.LastAddress.RowNumber == 1 && namedRange.RangeAddress.LastAddress.ColumnNumber == 1,
            "Pivot lookup name does not reference the allocated helper.");
        Ensure(pivotWorkbook.Worksheet("PivotData").DataValidations.Single().Value.TrimStart('=') == "PivotCodes",
            "Pivot lookup validation mismatch.");
    }
    using (var pivotArchive = new System.IO.Compression.ZipArchive(new MemoryStream(pivotBytes)))
    {
        System.Xml.Linq.XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        using var cacheStream = pivotArchive.Entries.Single(entry => entry.FullName.EndsWith("pivotCacheDefinition1.xml", StringComparison.Ordinal)).Open();
        using var recordsStream = pivotArchive.Entries.Single(entry => entry.FullName.EndsWith("pivotCacheRecords1.xml", StringComparison.Ordinal)).Open();
        using var tableStream = pivotArchive.Entries.Single(entry => entry.FullName.Contains("pivotTables/")
            && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).Open();
        var cache = System.Xml.Linq.XDocument.Load(cacheStream).Root;
        var records = System.Xml.Linq.XDocument.Load(recordsStream).Root.Elements(spreadsheet + "r").ToArray();
        var table = System.Xml.Linq.XDocument.Load(tableStream).Root;
        var source = cache.Element(spreadsheet + "cacheSource").Element(spreadsheet + "worksheetSource");
        Ensure((string)source.Attribute("sheet") == "PivotData" && (string)source.Attribute("ref") == "A1:B2",
            "Pivot cache source range mismatch.");
        var fields = cache.Element(spreadsheet + "cacheFields").Elements(spreadsheet + "cacheField").ToArray();
        Ensure(fields.Select(field => (string)field.Attribute("name")).SequenceEqual(new[] { "Code", "Count" })
            && records.Length == 1 && (string)records[0].Elements().Last().Attribute("v") == "7",
            "Pivot cache fields or numeric records mismatch.");
        var codeItem = records[0].Elements().First();
        var codeValue = codeItem.Name.LocalName == "x"
            ? (string)fields[0].Element(spreadsheet + "sharedItems").Elements()
                .ElementAt(int.Parse((string)codeItem.Attribute("v"), System.Globalization.CultureInfo.InvariantCulture))
                .Attribute("v")
            : (string)codeItem.Attribute("v");
        Ensure(codeValue == rows[0].Code && (string)table.Attribute("name") == "ConsumerPivot",
            "Pivot cache row text or table name mismatch.");
    }

    var formulaPivotRequest = ExcelExport.Workbook(workbook => workbook
        .AddSheet("PivotFormulaData", new[] { new ConsumerRow { Code = "=1+2", Count = 7 } })
        .PivotTable(new ExcelPivotTableDefinition
        {
            Name = "ConsumerFormulaPivot", SourceSheet = "PivotFormulaData",
            SourceRange = new ExcelRangeDefinition
            {
                StartRow = 0, StartColumn = 0, EndRow = 1, EndColumn = 1
            },
            TargetSheet = "FormulaSummary", RowField = "Code", ValueField = "Count"
        }));
    foreach (var asynchronous in new[] { false, true })
    {
        using var formulaOutput = new MemoryStream();
        var original = new byte[] { 7, 11, 19 };
        formulaOutput.Write(original, 0, original.Length);
        try
        {
            if (asynchronous)
                await closedXmlExporter.ExportAsync(formulaPivotRequest, formulaOutput);
            else
                closedXmlExporter.Export(formulaPivotRequest, formulaOutput);
            Ensure(false, "Pivot source formula must fail before evaluation.");
        }
        catch (BingOfficesConfigurationException exception)
        {
            Ensure(exception.Code == BingOfficesErrorCode.ConfigurationInvalid
                && exception.Stage == BingOfficesStage.Preflight
                && exception.SheetName == "PivotFormulaData"
                && exception.RowIndex == 2 && exception.ColumnIndex == 1
                && exception.Message.Contains("ConsumerFormulaPivot", StringComparison.Ordinal),
                "Pivot source formula failure lost its original coordinates.");
            Ensure(formulaOutput.ToArray().SequenceEqual(original) && formulaOutput.CanWrite,
                "Pivot source preflight failure changed or closed the caller stream.");
        }
    }

    var lookupRequest = ExcelExport.Workbook(workbook => workbook
        .LookupList("OrderStatus", new[] { "Open", "Closed" })
        .AddSheet("LookupOrders", new[] { new ConsumerLookupRow { Status = "Open" } },
            sheet => sheet.Protection(new ExcelSheetProtectionDefinition(new[]
                { new ExcelRangeDefinition { StartRow = 2, EndRow = 2, StartColumn = 0, EndColumn = 0 } }))
                .DataValidation(new ExcelDataValidationDefinition
            {
                Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 1, StartColumn = 0, EndColumn = 0 },
                Type = ExcelDataValidationType.List, ListSourceName = "OrderStatus"
            })));
    using var lookupOutput = new MemoryStream();
    directExporter.Export(lookupRequest, lookupOutput);
    using (var lookupWorkbook = new XSSFWorkbook(new MemoryStream(lookupOutput.ToArray())))
        Ensure(lookupWorkbook.NumberOfSheets == 2 && lookupWorkbook.GetSheetAt(1).GetRow(0)
            .GetCell(0).StringCellValue == "Open"
            && lookupWorkbook.GetSheetAt(0).Protect
            && !lookupWorkbook.GetSheetAt(0).GetRow(2).GetCell(0).CellStyle.IsLocked,
            "Packaged LookupList and protection workbook verification failed.");

    var entityLayout = ExcelEntity.Layout<ConsumerLookupRow>(builder => builder
        .Cell("Entity", "A1", item => item.Status)
        .Protection("Entity", new ExcelSheetProtectionDefinition(new[]
            { new ExcelRangeDefinition { StartRow = 1, EndRow = 1, StartColumn = 1, EndColumn = 1 } }))
        .LookupList("EntityStatus", new[] { "Open", "Closed" })
        .DataValidation("Entity", new ExcelDataValidationDefinition
        {
            Range = new ExcelRangeDefinition { StartRow = 0, EndRow = 0, StartColumn = 0, EndColumn = 0 },
            Type = ExcelDataValidationType.List, ListSourceName = "EntityStatus"
        }));
    using var entityOutput = new MemoryStream();
    directExporter.ExportEntity(new ConsumerLookupRow { Status = "Open" }, entityLayout, entityOutput);
    using var entityInput = new MemoryStream(entityOutput.ToArray());
    var entityResult = directImporter.ImportEntity(entityInput, entityLayout,
        new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail));
    Ensure(entityResult.IsSuccess && entityResult.Entity.Status == "Open",
        "Packaged Entity LookupList verification failed.");

    var batchLayout = ExcelEntity.Layout<ConsumerWorkbook>(builder => builder
        .Cell("Header", "B2", book => book.DocumentNumber)
        .ListRegion("Data", "A1", book => book.Rows));
    using var batchWorkbook = new XSSFWorkbook();
    batchWorkbook.CreateSheet("Header").CreateRow(1).CreateCell(1).SetCellValue("invalid");
    var batchSheet = batchWorkbook.CreateSheet("Data");
    var batchHeader = batchSheet.CreateRow(0);
    batchHeader.CreateCell(0).SetCellValue("Code");
    batchHeader.CreateCell(1).SetCellValue("Count");
    var batchData = batchSheet.CreateRow(1);
    batchData.CreateCell(0).SetCellValue("consumer-batch");
    batchData.CreateCell(1).SetCellValue(7);
    using var batchOutput = new MemoryStream();
    batchWorkbook.Write(batchOutput, true);
    var batchBytes = batchOutput.ToArray();
    foreach (var batchProvider in new[] { provider, closedXmlProvider })
    {
        var batchImporter = batchProvider.GetRequiredService<IExcelEntityBatchImporter>();
        var entityBatches = new List<ExcelEntityImportBatch<ConsumerWorkbook>>();
        using var batchInput = new MemoryStream(batchBytes);
        var batchSummary = batchImporter.ImportEntityBatches(batchInput, batchLayout,
            new ExcelEntityBatchImportOptions(1), entityBatches.Add);
        Ensure(!batchSummary.IsSuccess && batchSummary.Regions.Count == 1
            && batchSummary.SucceededRows == 1
            && batchSummary.Result.Entity.Rows.Count == 1
            && batchSummary.Result.Entity.Rows[0].Code == "consumer-batch"
            && batchSummary.Result.Entity.Rows[0].Count == 7
            && entityBatches.Count == 2
            && entityBatches[0].IsDiagnostic && entityBatches[0].Errors.Count > 0
            && !entityBatches[1].IsDiagnostic
            && entityBatches[1].SheetName == "Data"
            && entityBatches[1].Items.Count == 1,
            "Packaged Entity batch and diagnostic verification failed.");
    }

    var forwardLayout = ExcelEntity.Layout<ConsumerWorkbook>(builder => builder
        .ListRegion("Data", "A1", book => book.Rows));
    using var forwardWorkbook = new XSSFWorkbook();
    var forwardSheet = forwardWorkbook.CreateSheet("Data");
    var forwardHeader = forwardSheet.CreateRow(0);
    forwardHeader.CreateCell(0).SetCellValue("Code");
    forwardHeader.CreateCell(1).SetCellValue("Count");
    var forwardRow = forwardSheet.CreateRow(1);
    forwardRow.CreateCell(0).SetCellValue("consumer-forward");
    forwardRow.CreateCell(1).SetCellValue(9);
    using var forwardOutput = new MemoryStream();
    forwardWorkbook.Write(forwardOutput, true);
    var forwardBatches = new List<ExcelEntityImportBatch<ConsumerWorkbook>>();
    var forwardBatchImporter = provider.GetRequiredService<IExcelEntityBatchImporter>();
    using var forwardInput = new MemoryStream(forwardOutput.ToArray(), writable: false);
    var forwardSummary = forwardBatchImporter.ImportEntityBatches(forwardInput, forwardLayout,
        new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward,
            ExcelEntityBatchResultMode.SummaryOnly), forwardBatches.Add);
    Ensure(!forwardSummary.HasCompleteResult && forwardSummary.SucceededRows == 1
        && forwardSummary.Result.Entity.Rows.Count == 0
        && forwardBatches.Count == 1 && forwardBatches[0].Items.Count == 1
        && ((ConsumerRow)forwardBatches[0].Items[0]).Code == "consumer-forward"
        && ((ConsumerRow)forwardBatches[0].Items[0]).Count == 9,
        "Packaged Forward Entity batch verification failed.");

    var pageLayout = ExcelEntity.Layout<ConsumerWorkbook>(builder => builder.ListRegion("Data", "A1",
        book => book.Rows, region => region.PageBreak(1)
            .PageSubtotalNamed("ConsumerPage", "PAGE").Footer("END")));
    var pageDocument = new ConsumerWorkbook
    {
        Rows = new List<ConsumerRow>
        {
            new() { Code = "first", Count = 1 },
            new() { Code = "second", Count = 2 },
            new() { Code = "third", Count = 3 }
        }
    };
    foreach (var pageProvider in new[] { provider, closedXmlProvider })
    {
        using var pageOutput = new MemoryStream();
        ((IExcelEntityExporter)pageProvider.GetRequiredService<IExcelExporter>()).ExportEntity(
            pageDocument, pageLayout, pageOutput);
        using var pageWorkbook = new XSSFWorkbook(new MemoryStream(pageOutput.ToArray(), writable: false));
        var pageNames = pageWorkbook.GetAllNames().Where(name => name.NameName.StartsWith(
            "ConsumerPage_", StringComparison.OrdinalIgnoreCase)).ToArray();
        Ensure(pageNames.Length == 2 && pageNames.Any(name => name.NameName == "ConsumerPage_001")
            && pageNames.Any(name => name.NameName == "ConsumerPage_002"),
            "Packaged repeated page anchor verification failed.");
        using var pageInput = new MemoryStream(pageOutput.ToArray(), writable: false);
        var pageResult = ((IExcelEntityImporter)pageProvider.GetRequiredService<IExcelImporter>())
            .ImportEntity(pageInput, pageLayout);
        Ensure(pageResult.IsSuccess && pageResult.Entity.Rows.Count == 3,
            "Packaged repeated page anchor round-trip failed.");
    }

    var independentLayout = ExcelEntity.Layout<ConsumerWorkbook>(builder => builder
        .ListRegion("Data", "A2", book => book.Rows, region => region.PageBreak(1)
            .PageHeaderNamed("ConsumerHead", "A1:B2", "A1", "HEADER", header => header
                .Cell("B1", context => $"{context.PageNumber}/{context.PageCount}"))
            .PageFooterNamed("ConsumerSign", "SIGN", footer => footer.GapRows(1)
                .Cell("B1", items => items.Count).Cell("A2", "审核"),
                ExcelEntityPageAnchorStaleNamePolicy.RemoveObsolete).Footer("END")));
    foreach (var pageProvider in new[] { provider, closedXmlProvider })
    {
        using var independentOutput = new MemoryStream();
        await ((IExcelEntityExporter)pageProvider.GetRequiredService<IExcelExporter>())
            .ExportEntityAsync(pageDocument, independentLayout, independentOutput);
        using (var independentWorkbook = new XSSFWorkbook(new MemoryStream(independentOutput.ToArray())))
        {
            var sheet = independentWorkbook.GetSheet("Data");
            var expected = new[] { "HEADER|1/3", "Code|Count", "first|1", "|", "SIGN|1", "审核|",
                "HEADER|2/3", "Code|Count", "second|2", "|", "SIGN|1", "审核|",
                "HEADER|3/3", "Code|Count", "third|3", "|", "SIGN|1", "审核|", "END|" };
            Ensure(sheet.LastRowNum == 18 && expected.SequenceEqual(Enumerable.Range(0, 19)
                .Select(row => string.Join("|", Enumerable.Range(0, 2).Select(column =>
                    sheet.GetRow(row)?.GetCell(column)?.ToString() ?? "")))),
                "Packaged independent page layout content verification failed.");
            var names = independentWorkbook.GetAllNames();
            Ensure(names.Count == 6, "Packaged independent page name count failed.");
            for (var page = 0; page < 3; page++)
            {
                var headerName = names.Single(name => name.NameName == $"ConsumerHead_{page + 1:000}");
                var footerName = names.Single(name => name.NameName == $"ConsumerSign_{page + 1:000}");
                var headerReference = new NPOI.SS.Util.AreaReference(headerName.RefersToFormula,
                    NPOI.SS.SpreadsheetVersion.EXCEL2007);
                var footerReference = new NPOI.SS.Util.AreaReference(footerName.RefersToFormula,
                    NPOI.SS.SpreadsheetVersion.EXCEL2007);
                Ensure(headerReference.FirstCell.SheetName == "Data"
                    && headerReference.FirstCell.Row == page * 6 && headerReference.LastCell.Row == page * 6
                    && headerReference.FirstCell.Col == 0 && headerReference.LastCell.Col == 0
                    && footerReference.FirstCell.SheetName == "Data"
                    && footerReference.FirstCell.Row == page * 6 + 4 && footerReference.LastCell.Row == page * 6 + 4
                    && footerReference.FirstCell.Col == 0 && footerReference.LastCell.Col == 0,
                    $"Packaged independent page anchor coordinates failed: {headerName.RefersToFormula}; {footerName.RefersToFormula}.");
            }
        }
        independentOutput.Position = 0;
        var independentResult = await ((IExcelEntityImporter)pageProvider.GetRequiredService<IExcelImporter>())
            .ImportEntityAsync(independentOutput, independentLayout);
        Ensure(independentResult.IsSuccess && independentResult.Errors.Count == 0
            && pageDocument.Rows.Select(row => (row.Code, row.Count)).SequenceEqual(
                independentResult.Entity.Rows.Select(row => (row.Code, row.Count))),
            "Packaged independent page round-trip failed.");
    }

    var printLayout = ExcelEntity.Layout<ConsumerWorkbook>(builder => builder
        .PrintTitleRows("Data", "1:1")
        .ListRegion("Data", "A2", book => book.Rows));
    var headerLayout = ExcelEntity.Layout<ConsumerWorkbook>(builder => builder
        .ListRegion("Data", "A2", book => book.Rows, region => region
            .PageBreak(1).PageHeader("A1:B2", "A1", "HEADER", header => header
                .Cell("B1", (ExcelEntityPageHeaderContext<ConsumerRow> context) =>
                    $"{context.PageNumber}/{context.PageCount}"))));
    foreach (var pageProvider in new[] { provider, closedXmlProvider })
    {
        var exporter = (IExcelEntityExporter)pageProvider.GetRequiredService<IExcelExporter>();
        var importer = (IExcelEntityImporter)pageProvider.GetRequiredService<IExcelImporter>();
        using var printOutput = new MemoryStream();
        exporter.ExportEntity(pageDocument, printLayout, printOutput);
        using (var printWorkbook = new XSSFWorkbook(new MemoryStream(printOutput.ToArray())))
            Ensure(printWorkbook.GetSheet("Data").RepeatingRows?.FirstRow == 0,
                "Packaged Entity print title verification failed.");

        using var headerOutput = new MemoryStream();
        exporter.ExportEntity(pageDocument, headerLayout, headerOutput);
        using (var headerWorkbook = new XSSFWorkbook(new MemoryStream(headerOutput.ToArray())))
        {
            var sheet = headerWorkbook.GetSheet("Data");
            Ensure(sheet.GetRow(0).GetCell(0).StringCellValue == "HEADER"
                && sheet.GetRow(3).GetCell(0).StringCellValue == "HEADER"
                && sheet.GetRow(6).GetCell(0).StringCellValue == "HEADER"
                && sheet.GetRow(3).GetCell(1).StringCellValue == "2/3",
                "Packaged Entity repeated header verification failed.");
        }
        using var headerInput = new MemoryStream(headerOutput.ToArray(), writable: false);
        var headerResult = importer.ImportEntity(headerInput, headerLayout);
        Ensure(headerResult.IsSuccess && headerResult.Entity.Rows.Count == 3,
            "Packaged Entity repeated header round-trip failed.");
    }

    var cascadeRequest = ExcelExport.Workbook(book => book
        .CascadeList("Categories", new Dictionary<string, IReadOnlyList<string>>
        {
            ["Open"] = new[] { "Ready", "Pending" }
        })
        .AddSheet("Cascade", new[] { new ConsumerLookupRow { Status = "Open" } }, sheet =>
            sheet.DataValidation(new ExcelDataValidationDefinition
            {
                Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 5, StartColumn = 1, EndColumn = 1 },
                Type = ExcelDataValidationType.List,
                CascadeSourceName = "Categories", ParentColumnIndex = 0
            })));
    using var cascadeOutput = new MemoryStream();
    directExporter.Export(cascadeRequest, cascadeOutput);
    using (var cascadeBook = new XSSFWorkbook(new MemoryStream(cascadeOutput.ToArray())))
        Ensure(cascadeBook.NumberOfSheets == 2
            && cascadeBook.GetSheetAt(1).GetRow(0).GetCell(2).StringCellValue == "Ready"
            && cascadeBook.GetSheet("Cascade").GetDataValidations()[0].ValidationConstraint.Formula1
                .StartsWith("INDIRECT(VLOOKUP($A2,", StringComparison.Ordinal),
            "Packaged cascade dictionary verification failed.");

    using var readOnlyProvider = new ServiceCollection()
        .AddBingOfficesExcelDataReader()
        .BuildServiceProvider();
    var readOnlyImporter = readOnlyProvider.GetRequiredService<IExcelImporter>();
    var readOnlyResult = await ExcelStreamExtensions.ImportFromBytesAsync(
        readOnlyImporter, excelAsyncBytes, importRequest);
    var workbookBatchImporter = readOnlyProvider.GetRequiredService<IExcelWorkbookBatchImporter>();
    var workbookBatches = new List<ExcelWorkbookImportBatch>();
    var workbookBatchResult = workbookBatchImporter.ImportWorkbookBatches(
        new MemoryStream(excelAsyncBytes), importRequest, workbookBatches.Add);
    await Bing.Offices.PackageConsumer.ExcelDataReaderFailurePackageContract.VerifyAsync(
        readOnlyProvider.GetRequiredService<IExcelBatchImporter>(), excelAsyncBytes);

    using var streamingOutput = new MemoryStream();
    await new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(workbookRequest, streamingOutput);
    using var streamedWorkbook = new XSSFWorkbook(new MemoryStream(streamingOutput.ToArray()));
    Ensure(streamedWorkbook.NumberOfSheets == 1
        && streamedWorkbook.GetSheetAt(0).GetRow(1).GetCell(0).StringCellValue == rows[0].Code
        && streamedWorkbook.GetSheetAt(0).GetRow(1).GetCell(1).NumericCellValue == rows[0].Count,
        "SpreadCheetah packaged output content verification failed.");

    using var extensionWorkbook = new XSSFWorkbook();
    var extensionSheet = extensionWorkbook.CreateSheet("Extensions");
    extensionSheet.CreateRow(0).Value(0, "extension");

    Ensure(csvSyncResult.Items.Count == 1 && csvAsyncResult.Items.Count == 1,
        "CSV byte sync/async verification failed.");
    Ensure(csvFileResult.Items.Count == 1 && csvAsyncFileResult.Items.Count == 1,
        "CSV file sync/async verification failed.");
    Ensure(excelSyncResult.Workbook.Rows.Count == 1 && excelAsyncResult.Workbook.Rows.Count == 1,
        "Excel byte sync/async verification failed.");
    Ensure(excelFileResult.Workbook.Rows.Count == 1 && excelAsyncFileResult.Workbook.Rows.Count == 1,
        "Excel file sync/async verification failed.");
    Ensure(directResult.Workbook.Rows.Count == 1 && directBytes.Length > 0,
        "Direct NPOI provider verification failed.");
    Ensure(miniResult.Workbook.Rows.Count == 1 && miniBytes.Length > 0,
        "MiniExcel provider verification failed.");
    Ensure(closedXmlResult.Workbook.Rows.Count == 1 && closedXmlBytes.Length > 0,
        "ClosedXML provider verification failed.");
    Ensure(readOnlyResult.Workbook.Rows.Count == 1,
        "ExcelDataReader provider verification failed.");
    Ensure(workbookBatchResult.IsSuccess && workbookBatches.Count == 1
        && workbookBatches[0].Items.Count == 1,
        "ExcelDataReader workbook batch verification failed.");
    Ensure(extensionWorkbook.GetExcelFormat() == ExcelFormat.Xlsx
        && extensionSheet.GetRow(0).GetCell(0).GetStringValue() == "extension",
        "NPOI extension verification failed.");

    Console.WriteLine($"package-consumer-ok tfm={System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription} packages=Bing.Offices.Npoi+Bing.Offices.MiniExcel+Bing.Offices.ClosedXml+Bing.Offices.ExcelDataReader+Bing.Offices.SpreadCheetah+Bing.Offices.AsposeCells/{packageVersion} csvBytes={csvAsyncBytes.Length} excelBytes={excelAsyncBytes.Length} miniExcelBytes={miniBytes.Length} closedXmlBytes={closedXmlBytes.Length} excelDataReaderRows={readOnlyResult.Workbook.Rows.Count} npoiExtensions=ok");
}
finally
{
    if (Directory.Exists(directory))
        Directory.Delete(directory, true);
}

static void Ensure(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}

/// <summary>
/// 表示消费者程序用于往返验证的单行数据。
/// </summary>
public sealed class ConsumerRow
{
    /// <summary>
    /// 获取或设置行的编码。
    /// </summary>
    public string Code { get; set; }

    /// <summary>
    /// 获取或设置行中的数量。
    /// </summary>
    public int Count { get; set; }
}

/// <summary>
/// 表示消费者程序用于 Excel 往返验证的工作簿。
/// </summary>
public sealed class ConsumerWorkbook
{
    /// <summary>
    /// 获取或设置单据编号。
    /// </summary>
    public int DocumentNumber { get; set; }

    /// <summary>
    /// 获取工作簿中的数据行集合。
    /// </summary>
    public List<ConsumerRow> Rows { get; set; } = new();
}

/// <summary>
/// 表示字典模板消费者的实体。
/// </summary>
public sealed class ConsumerLookupRow
{
    /// <summary>
    /// 获取或设置订单状态。
    /// </summary>
    public string Status { get; set; }
}
