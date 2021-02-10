using System.Collections.Generic;

namespace Export.Common.Utils.Excel
{
    public interface IConfiguracionTabla
    {
        IList<Nota> Cabeceras { get; set; }
        IList<Nota> Pies { get; set; }
        PosicionTabla PosicionInicial { get; set; }
        bool ShowAutoFilter { get; set; }
    }
}