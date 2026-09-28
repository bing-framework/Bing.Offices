using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using Bing.Offices.Exceptions;

namespace Bing.Offices.Configurations;

/// <summary>
/// 映射配置加载器的默认服务实现。
/// </summary>
internal sealed class DefaultExcelMappingConfigurationLoader : IExcelMappingConfigurationLoader
{
    /// <summary>
    /// 向注册的观察器转发配置加载异常。
    /// </summary>
    private readonly BingOfficesExceptionDispatcher _exceptionDispatcher;

    /// <summary>
    /// 初始化一个 <see cref="DefaultExcelMappingConfigurationLoader" /> 类型的实例。
    /// </summary>
    /// <param name="exceptionObservers">接收配置加载异常的可选观察器集合。</param>
    public DefaultExcelMappingConfigurationLoader(IEnumerable<IBingOfficesExceptionObserver> exceptionObservers = null)
    {
        _exceptionDispatcher = new BingOfficesExceptionDispatcher(exceptionObservers);
    }

    /// <inheritdoc />
    public ExcelMappingDocument FromJsonDocument(string json) =>
        Execute(() => ExcelMappingConfigurationLoader.FromJsonDocument(json));

    /// <inheritdoc />
    public ExcelMappingDocument FromJsonDocument(Stream source) =>
        Execute(() => ExcelMappingConfigurationLoader.FromJsonDocument(source));

    /// <inheritdoc />
    public ExcelMappingDocument FromXmlDocument(string xml) =>
        Execute(() => ExcelMappingConfigurationLoader.FromXmlDocument(xml));

    /// <inheritdoc />
    public ExcelMappingDocument FromXmlDocument(Stream source) =>
        Execute(() => ExcelMappingConfigurationLoader.FromXmlDocument(source));

    /// <summary>
    /// 执行默认加载操作，并将配置异常通知观察器。
    /// </summary>
    /// <param name="load">待执行的映射文档加载操作。</param>
    /// <returns>加载后的映射文档。</returns>
    private ExcelMappingDocument Execute(Func<ExcelMappingDocument> load)
    {
        try
        {
            return load();
        }
        catch (BingOfficesException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
    }
}
