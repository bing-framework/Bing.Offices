using System;
using Bing.Offices.Tests.Models;
using Xunit;

namespace Bing.Offices.Tests
{
    public class FluentTest
    {
        [Fact]
        public void Test_FluentApi()
        {
            var fluent = FluentSettings.For<FluentSample>();
            fluent.Property(x => x.FluentTitle).HasColumnTitle("Fluent标题");
            Console.WriteLine("");
        }
    }
}
