using System.Collections;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using BenchmarkDotNet.Attributes;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Csv;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Providers;
using Microsoft.Extensions.DependencyInjection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Benchmarks;

/// <summary>分别测量真实表头绑定和一行数据的公共导入路径；输入及内部调用委托在计时前准备。</summary>
[MemoryDiagnoser]
public class HeaderBindingBenchmarks
{
    /// <summary>格式与列数；XLS 不超过其 256 列格式上限。</summary>
    [ParamsSource(nameof(Cases))]
    public string Case { get; set; } = "csv-16";

    /// <summary>有限的普通表和宽表合成数据集。</summary>
    public IEnumerable<string> Cases => new[]
    {
        "csv-16", "csv-64", "csv-256", "csv-1024",
        "xls-16", "xls-64", "xls-256",
        // 请求映射最多允许 1000 个动态定义；两列固定字段包含在总列数内。
        "xlsx-16", "xlsx-64", "xlsx-256", "xlsx-1000"
    };

    private ServiceProvider _services = null!;
    private ICsvImporter _csvImporter = null!;
    private IExcelImporter _excelImporter = null!;
    private CsvImportOptions<HeaderRow> _csvOptions = null!;
    private ExcelWorkbookImportRequest<HeaderWorkbook> _request = null!;
    private Func<object> _bind = null!;
    private IWorkbook? _workbook;
    private byte[] _payload = Array.Empty<byte>();
    private string _format = string.Empty;
    private int _columnCount;

    /// <summary>创建相同语义的输入与映射，编译内部调用委托，并验证两次完整导入结果。</summary>
    [GlobalSetup]
    public void Setup()
    {
        var parts = Case.Split('-');
        _format = parts[0];
        _columnCount = int.Parse(parts[1], System.Globalization.CultureInfo.InvariantCulture);
        _services = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        _csvImporter = _services.GetRequiredService<ICsvImporter>();
        _excelImporter = _services.GetRequiredService<IExcelImporter>();
        var configuration = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(HeaderRow.Name), Title = "Display", Aliases = new List<string> { "LegacyName" } }
            },
            DynamicColumns = Enumerable.Range(0, _columnCount - 2).Select(index =>
                new ExcelMappingDynamicColumnConfiguration
                {
                    Key = $"k{index}", Title = $"Column{index}", Order = index,
                    Aliases = new List<string> { $"Alias{index}" }, DataTypeName = "string"
                }).ToList()
        };
        var headers = new[] { "LegacyName", "Tail" }.Concat(configuration.DynamicColumns.Select((column, index) =>
            index % 2 == 0 ? column.Title : column.Aliases[0])).ToArray();
        var values = new[] { "head", "tail" }.Concat(Enumerable.Range(0, _columnCount - 2).Select(index => $"v{index}")).ToArray();
        var plan = _services.GetRequiredService<IExcelMappingPlanFactory>()
            .Create<HeaderRow>(new ExcelMappingDocument(), configuration, MappingDirection.Import);
        _csvOptions = new CsvImportOptions<HeaderRow> { MappingConfiguration = configuration, MaxColumns = _columnCount };
        _request = ExcelImport.Workbook<HeaderWorkbook>(builder =>
            builder.Sheet("Data", root => root.Rows, sheet => sheet.Mapping(configuration).MaxReadColumns(_columnCount)));

        if (_format == "csv")
        {
            _payload = Encoding.UTF8.GetBytes(string.Join(",", headers) + "\r\n" + string.Join(",", values) + "\r\n");
            _bind = CreateCsvBinding(headers, plan);
        }
        else
        {
            _workbook = _format == "xls" ? new HSSFWorkbook() : new XSSFWorkbook();
            var sheet = _workbook.CreateSheet("Data");
            var header = sheet.CreateRow(0);
            var row = sheet.CreateRow(1);
            for (var index = 0; index < _columnCount; index++)
            {
                header.CreateCell(index).SetCellValue(headers[index]);
                row.CreateCell(index).SetCellValue(values[index]);
            }
            using var stream = new MemoryStream();
            _workbook.Write(stream, true);
            _payload = stream.ToArray();
            _bind = CreateNpoiBinding(header, plan);
        }

        for (var iteration = 0; iteration < 2; iteration++)
        {
            if (((IEnumerable)_bind()).Cast<object>().Count() != _columnCount)
                throw new InvalidOperationException("Binding did not preserve every column.");
            Validate(ImportRow());
        }
    }

    /// <summary>仅调用实际生产绑定方法；包含其每次调用需要的临时集合分配。</summary>
    [Benchmark]
    public object Bind() => _bind();

    /// <summary>完整公共导入路径，包含解析／DOM、计划查询、绑定和一行转换。</summary>
    [Benchmark]
    public HeaderRow ImportOneRow() => ImportRow();

    /// <summary>释放本轮工作簿和依赖注入容器。</summary>
    [GlobalCleanup]
    public void Cleanup()
    {
        _workbook?.Dispose();
        _services.Dispose();
    }

    private HeaderRow ImportRow()
    {
        using var source = new MemoryStream(_payload, writable: false);
        if (_format == "csv")
        {
            var result = _csvImporter.Import(source, _csvOptions);
            if (result.Errors.Count != 0 || result.Items.Count != 1)
                throw new InvalidOperationException("CSV import failed.");
            return result.Items[0];
        }
        var workbookResult = _excelImporter.Import(source, _request);
        if (workbookResult.Errors.Count != 0 || workbookResult.Workbook.Rows.Count != 1)
            throw new InvalidOperationException("Excel import failed.");
        return workbookResult.Workbook.Rows[0];
    }

    private void Validate(HeaderRow row)
    {
        if (row.Name != "head" || row.Tail != "tail" || row.Values.Count != _columnCount - 2)
            throw new InvalidOperationException("Imported fixed values or dynamic count changed.");
        for (var index = 0; index < _columnCount - 2; index++)
            if (!row.Values.TryGetValue($"k{index}", out var value) || !Equals(value, $"v{index}"))
                throw new InvalidOperationException($"Dynamic value changed at {index}.");
    }

    private static Func<object> CreateCsvBinding(string[] headers, IExcelMappingPlan plan)
    {
        var assembly = typeof(ExcelMappingPlanFactoryProvider).Assembly;
        var bindingType = assembly.GetType("Bing.Offices.Csv.CsvPropertyBinding", throwOnError: true)!;
        var create = bindingType.GetMethod("Create", BindingFlags.Static | BindingFlags.NonPublic)!
            .MakeGenericMethod(typeof(HeaderRow));
        Array BindProperties(bool dynamic)
        {
            var columns = plan.Columns.Where(column => !column.Ignored && column.IsDynamicColumn == dynamic).ToArray();
            var bindings = Array.CreateInstance(bindingType, columns.Length);
            for (var index = 0; index < columns.Length; index++)
                bindings.SetValue(create.Invoke(null, new object[] { columns[index] }), index);
            return bindings;
        }
        var method = assembly.GetType("Bing.Offices.Csv.CsvHeaderBinder", throwOnError: true)!
            .GetMethods(BindingFlags.Static | BindingFlags.Public)
            .Single(candidate => candidate.Name == "Bind" && candidate.GetParameters()[0].ParameterType == typeof(IReadOnlyList<string>));
        return CompileCall(method, headers, BindProperties(false), BindProperties(true), plan.DynamicColumns, true, null);
    }

    private Func<object> CreateNpoiBinding(IRow header, IExcelMappingPlan plan)
    {
        var assembly = _excelImporter.GetType().Assembly;
        var optionsType = assembly.GetType("Bing.Offices.Imports.ExcelImportExecutionOptions`1", throwOnError: true)!
            .MakeGenericType(typeof(HeaderRow));
        var options = Activator.CreateInstance(optionsType)!;
        optionsType.GetProperty("MappingPlan", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(options, plan);
        optionsType.GetProperty("MaxReadColumns", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(options, _columnCount);
        var method = assembly.GetType("Bing.Offices.Imports.NpoiImportSheetExecutor", throwOnError: true)!
            .GetMethod("CreateColumns", BindingFlags.Static | BindingFlags.NonPublic)!.MakeGenericMethod(typeof(HeaderRow));
        return CompileCall(method, header, options);
    }

    /// <summary>反射仅用于 Setup；计时期间通过已编译委托调用真实内部实现。</summary>
    private static Func<object> CompileCall(MethodInfo method, params object?[] arguments)
    {
        var parameters = method.GetParameters();
        var call = Expression.Call(method, arguments.Select((argument, index) =>
            Expression.Constant(argument, parameters[index].ParameterType)));
        return Expression.Lambda<Func<object>>(Expression.Convert(call, typeof(object))).Compile();
    }

    /// <summary>公共入口使用的合成行，固定字段与动态定义同时参与绑定。</summary>
    public sealed class HeaderRow
    {
        public string Name { get; set; } = string.Empty;
        public string Tail { get; set; } = string.Empty;
        [DynamicColumn]
        public IDictionary<string, object> Values { get; set; } = new Dictionary<string, object>();
    }

    /// <summary>单 Sheet 导入结果容器。</summary>
    public sealed class HeaderWorkbook
    {
        public List<HeaderRow> Rows { get; } = new();
    }
}
