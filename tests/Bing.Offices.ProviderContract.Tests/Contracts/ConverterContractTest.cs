using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.ProviderContract.Tests.Runner;
using Bing.Offices.Testing.Requests;
using Bing.Offices.Testing.Snapshots;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证命名值转换器在三个 Provider 上使用同一公共请求。
/// </summary>
public sealed class ConverterContractTest
{
    /// <summary>
    /// 获取参与契约测试的 Provider 名称。
    /// </summary>
    public static IEnumerable<object[]> Providers => ProviderDrivers.All.Select(driver => new object[] { driver.Name });

    /// <summary>
    /// 导入转换器应产生同一独立快照。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task NamedConverterContract_ShouldMatchIndependentSnapshot(string provider)
    {
        var converter = new UpperStringConverter();
        var execution = await ProviderContractRunner.RoundTripAsync(
            provider,
            "Converter",
            ProviderDrivers.Get(provider).CreateExporter(new[] { converter }),
            ProviderDrivers.Get(provider).CreateImporter(new[] { converter }),
            ContractRequests.MappingExportWithConverter(),
            ContractRequests.MappingImportWithConverter());

        ProviderContractAssertions.HasNoException(execution);
        ProviderContractAssertions.HasResult(execution);
        Assert.True(execution.Result.IsSuccess,
            execution.FormatAssertionFailure($"converter import failed: {string.Join(";", execution.Result.Errors.Select(error => error.Message))}"));
        Assert.True(new[] { "ALICE", "BOB" }.SequenceEqual(execution.Result.Workbook.Rows.Select(row => row.Name)),
            execution.FormatAssertionFailure("converted values differ from the independent snapshot"));
        Assert.True(converter.ImportColumns.Any(column => column > 0),
            execution.FormatAssertionFailure("converter did not receive a physical import column"));
    }

    /// <summary>
    /// 将导入字符串转为大写并记录列索引的测试转换器。
    /// </summary>
    private sealed class UpperStringConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "contract-upper";
        /// <summary>
        /// 获取导入时接收的列索引。
        /// </summary>
        public List<int> ImportColumns { get; } = new();
        /// <inheritdoc />
        public bool CanConvert(System.Type propertyType) => propertyType == typeof(string);
        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            ImportColumns.Add(context.ColumnIndex);
            value = context.Value?.ToString()?.ToUpperInvariant();
            return true;
        }
        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = context.Value;
            return true;
        }
    }
}
