using Bing.Offices.Imports;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.PackageConsumer;

internal static class ExcelDataReaderFailurePackageContract
{
    public static async Task VerifyAsync(IExcelBatchImporter importer, byte[] validWorkbookBytes)
    {
        var invalidWorkbook = CreateInvalidWorkbook(validWorkbookBytes);
        VerifySync(importer, invalidWorkbook.Bytes, invalidWorkbook.Code);
        await VerifyAsyncCore(importer, invalidWorkbook.Bytes, invalidWorkbook.Code)
            .ConfigureAwait(false);
    }

    private static (byte[] Bytes, string Code) CreateInvalidWorkbook(byte[] validWorkbookBytes)
    {
        using var source = new MemoryStream(validWorkbookBytes, writable: false);
        using var workbook = new XSSFWorkbook(source);
        var sheet = workbook.GetSheet("Data");
        var row = sheet?.GetRow(1);
        var codeCell = row?.GetCell(0);
        var countCell = row?.GetCell(1);
        Ensure(codeCell != null && countCell != null,
            "Packaged failure workbook source row was not found.");
        var code = codeCell.StringCellValue;
        countCell.SetCellValue("invalid-count");

        using var output = new MemoryStream();
        workbook.Write(output);
        return (output.ToArray(), code);
    }

    private static void VerifySync(IExcelBatchImporter importer, byte[] invalidWorkbookBytes,
        string expectedCode)
    {
        using var source = new MemoryStream(invalidWorkbookBytes, writable: false);
        using var destination = new MemoryStream();
        var batches = new List<ExcelImportBatch<global::ConsumerRow>>();
        var summary = importer.ImportBatches(source, CreateRequest(destination), batches.Add);

        VerifySummary(summary, batches, "sync");
        Ensure(source.CanRead && destination.CanRead && destination.Length > 0,
            "Packaged sync failure import changed stream ownership.");
        VerifyWorkbook(destination, "sync", expectedCode);
    }

    private static async Task VerifyAsyncCore(IExcelBatchImporter importer,
        byte[] invalidWorkbookBytes, string expectedCode)
    {
        using var source = new MemoryStream(invalidWorkbookBytes, writable: false);
        using var destination = new MemoryStream();
        var batches = new List<ExcelImportBatch<global::ConsumerRow>>();
        var summary = await importer.ImportBatchesAsync(source, CreateRequest(destination),
            (batch, _) =>
            {
                batches.Add(batch);
                return Task.CompletedTask;
            }).ConfigureAwait(false);

        VerifySummary(summary, batches, "async");
        Ensure(source.CanRead && destination.CanRead && destination.Length > 0,
            "Packaged async failure import changed stream ownership.");
        VerifyWorkbook(destination, "async", expectedCode);
    }

    private static ExcelBatchImportRequest<global::ConsumerRow> CreateRequest(Stream destination)
    {
        return new ExcelBatchImportRequest<global::ConsumerRow>("Data")
        {
            BatchSize = 1,
            ValidationFailureMode = ExcelValidationFailureMode.Continue,
            FailureOptions = new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            }
        };
    }

    private static void VerifySummary(ExcelBatchImportSummary summary,
        IReadOnlyList<ExcelImportBatch<global::ConsumerRow>> batches, string mode)
    {
        Ensure(!summary.IsSuccess && summary.RowsRead == 1 && summary.SucceededRows == 0
            && summary.FailedRows == 1 && summary.ErrorCount == 1,
            $"Packaged {mode} failure summary was incomplete.");
        Ensure(batches.Count == 1 && batches[0].Items.Count == 0
            && batches[0].Errors.Count == 1
            && batches[0].Errors[0].Code == ExcelImportErrorCode.ValueConversion
            && batches[0].Errors[0].PropertyName == nameof(global::ConsumerRow.Count),
            $"Packaged {mode} failure batch was incomplete.");
    }

    private static void VerifyWorkbook(MemoryStream destination, string mode,
        string expectedCode)
    {
        destination.Position = 0;
        using var workbook = new XSSFWorkbook(destination);
        var data = workbook.GetSheet("Data");
        var errors = workbook.GetSheet("__BingErrors");
        Ensure(data != null && errors != null,
            $"Packaged {mode} failure workbook is missing a required sheet.");
        Ensure(data.GetRow(0).GetCell(2).StringCellValue == "__Errors"
            && data.GetRow(1).GetCell(0).StringCellValue == expectedCode
            && data.GetRow(1).GetCell(1).StringCellValue == "invalid-count"
            && !string.IsNullOrWhiteSpace(data.GetRow(1).GetCell(2).StringCellValue),
            $"Packaged {mode} failure row did not preserve its source values.");
        Ensure(errors.LastRowNum == 1
            && errors.GetRow(0).GetCell(0).StringCellValue == "Sheet"
            && errors.GetRow(0).GetCell(1).StringCellValue == "Row"
            && errors.GetRow(1).GetCell(0).StringCellValue == "Data"
            && errors.GetRow(1).GetCell(1).NumericCellValue == 2
            && errors.GetRow(1).GetCell(3).StringCellValue
                == nameof(ExcelImportErrorCode.ValueConversion),
            $"Packaged {mode} failure workbook error coordinates were incomplete.");
    }

    private static void Ensure(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }
}
