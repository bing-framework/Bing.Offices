using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models.Bugs
{
    public class Issue1
    {
        [ColumnName("值")]
        public decimal DecimalValue { get; set; }
    }
}
