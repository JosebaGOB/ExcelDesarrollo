using ClosedXML.Excel;

namespace ConsoleApplication18.Excel
{
    /// <summary>
    /// Definirá la posición inicial de la tabla
    /// Por defecto será 0,0 dentro de la página
    /// </summary>
    public class PosicionTabla
    {
        public int Columna { get; set; }
        public int Fila { get; set; }

        public PosicionTabla() : this(1, 1)
        {
        }

        public PosicionTabla(int columna, int fila)
        {
            Columna = columna;
            Fila = fila;
        }
    }
}