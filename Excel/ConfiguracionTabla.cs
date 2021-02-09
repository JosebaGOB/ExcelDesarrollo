using System.Collections.Generic;

namespace ConsoleApplication18.Excel
{
    public interface IConfiguracionTabla
    {
         IList<Nota> Cabeceras { get; set; }
         IList<Nota> Pies { get; set; }
         PosicionTabla PosicionInicial { get; set; }
         bool ShowAutoFilter { get; set; }
    }

    public class ConfiguracionTabla : IConfiguracionTabla
    {
        public Estilo EstiloTabla { get; set; }
        public IList<Nota> Cabeceras { get; set; }
        public IList<Nota> Pies { get; set; }
        public PosicionTabla PosicionInicial { get; set; }
        public bool ShowAutoFilter { get; set; }

        public ConfiguracionTabla()
        {
            Cabeceras = new List<Nota>();
            Pies = new List<Nota>();
            PosicionInicial = new PosicionTabla();
            EstiloTabla = new Estilo();
            ShowAutoFilter = false;
        }

    }
}