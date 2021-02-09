using System.Collections.Generic;
using ClosedXML.Excel;

namespace ConsoleApplication18.Excel
{
    public class ConfiguracionTablaTheme : IConfiguracionTabla
    {
        public XLTableTheme ThemeTabla { get; set; }
        public bool ShowAutoFilter { get; set; }
        public IList<Nota> Cabeceras { get; set; }
        public IList<Nota> Pies { get; set; }
        public PosicionTabla PosicionInicial { get; set; }

        public ConfiguracionTablaTheme()
        {
            Cabeceras = new List<Nota>();
            Pies = new List<Nota>();
            PosicionInicial = new PosicionTabla();
            ShowAutoFilter = false;
            ThemeTabla = XLTableTheme.None;
        }
    }
}