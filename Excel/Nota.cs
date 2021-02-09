using ClosedXML.Excel;

namespace ConsoleApplication18.Excel
{
    /// <summary>
    /// Introduce en la zona superior o inferior de la tabla
    /// un texto con un estilo que se puede modificar manualmente
    /// El que se sitúe por encima o debajo dependerá de cómo
    /// esté definido dentro de la configuración de la tabla
    /// si como "nota" o como "pie".
    /// </summary>
    public class Nota
    {
        public Estilo Estilo { get; set; }

        // Por defecto las notas tendrán
        // tantas celdas de ancho como
        // columnas la tabla
        public int AnchoCeldas { get; set; }

        public string Texto { get; set; }

        public Nota() { }

        public Nota(string texto): this(texto, null)
        {

        }

        public Nota(string texto, Estilo estilo)
        {
            Texto = texto;
            Estilo = estilo;
        }


    }
}