using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models
{
    [Sheet(Index = 1, Name = "Test", Title = "测试一下", StartRowIndex = 1)]
    public class FluentSample
    {
        [Column("标识", Index = 1)]
        public int Id { get; set; }

        [Column("名称", Index = 2)]
        public string Name { get; set; }

        [Column("内容", Index = 3)]
        public string Content { get; set; }

        [Column("宽度", Index = 4, Width = 300)]
        public int Width { get; set; }

        [Column(IsIgnored = true)]
        public bool IsIgnored { get; set; }

        [Column("格式化", Index = 5, Formatter = "XXX")]
        public string Formatter { get; set; }

        [Column(6)]
        public string NoTitle { get; set; }

        [Column(7, Title = "Attribute标题")]
        public string FluentTitle { get; set; }
    }
}
