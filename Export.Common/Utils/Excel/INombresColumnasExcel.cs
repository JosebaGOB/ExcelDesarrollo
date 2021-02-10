using System.Collections.Generic;

namespace Export.Common.Utils.Excel
{
    public interface INombresColumnasExcel
    {
        Dictionary<string, string> NombresColumnas { get; }
    }
}