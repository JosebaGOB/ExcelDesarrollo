using System.Collections.Generic;

namespace Export.Common.Utils.Excel
{
    public class ConfiguracionTablaManual : IConfiguracionTabla
    {
        public Estilo EstiloTabla { get; set; }
        public IList<Nota> Cabeceras { get; set; }
        public IList<Nota> Pies { get; set; }
        public PosicionTabla PosicionInicial { get; set; }
        public bool ShowAutoFilter { get; set; }

        public ConfiguracionTablaManual()
        {
            Cabeceras = new List<Nota>();
            Pies = new List<Nota>();
            PosicionInicial = new PosicionTabla();
            EstiloTabla = new Estilo();
            ShowAutoFilter = false;
        }

    }
}