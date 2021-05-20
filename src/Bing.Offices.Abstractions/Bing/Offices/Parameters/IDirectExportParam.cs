using System;
using System.Collections.Generic;
using System.Text;

namespace Bing.Offices.Parameters
{
    public interface IDirectExportParam : IExportParam
    {
        string SheetName { get; set; }

    }
}
