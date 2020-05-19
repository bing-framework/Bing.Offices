using System.Collections.Generic;
using System.Reflection;
using Bing.Helpers;
using Bing.Offices.Excel;
using Bing.Offices.Excel.Mappings;
using Bing.Utils.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Bing.Offices.Mappings.Tests
{
    /// <summary>
    /// 映射测试
    /// </summary>
    public class MapTest : TestBase
    {
        /// <summary>
        /// 上下文
        /// </summary>
        internal ExcelContext Context { get; }

        /// <summary>
        /// 初始化一个<see cref="TestBase"/>类型的实例
        /// </summary>
        /// <param name="output">输出</param>
        public MapTest(ITestOutputHelper output) : base(output)
        {
            Context = new ExcelContext();
            Mapping();
        }

        /// <summary>
        /// 获取程序集
        /// </summary>
        private Assembly[] GetAssemblies() => new[] {GetType().Assembly};

        /// <summary>
        /// 从程序集获取映射配置列表
        /// </summary>
        private IEnumerable<IExcelExportMap> GetMapsFromAssemblies()
        {
            var result = new List<IExcelExportMap>();
            foreach (var assembly in GetAssemblies())
                result.AddRange(GetMapInstances(assembly));
            return result;
        }

        /// <summary>
        /// 获取映射实例列表
        /// </summary>
        /// <param name="assembly">程序集</param>
        private IEnumerable<IExcelExportMap> GetMapInstances(Assembly assembly) => Reflection.GetInstancesByInterface<IExcelExportMap>(assembly);

        /// <summary>
        /// 配置映射
        /// </summary>
        private void Mapping()
        {
            foreach (var mapper in GetMapsFromAssemblies())
                mapper.Map(Context);
        }

        [Fact]
        public void Test_InitMap()
        {
            
        }
    }
}
