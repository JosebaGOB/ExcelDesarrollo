using System.Collections.Generic;
using System.Data;
using ClosedXML.Excel;
using ConsoleApplication18.Excel;


namespace ConsoleApplication18
{
    public class TablaExcel
    {

        public XLWorkbook CrearDocumento<T>(List<T> datos, string tituloPagina,
            ConfiguracionTabla configuracionTabla = null)
        {
            var dataTable = datos.ToDataTable();

            return CrearDocumento(dataTable, tituloPagina, configuracionTabla);
        }
        
        
        public XLWorkbook CrearDocumento(DataTable datos, string tituloPagina,
            ConfiguracionTabla configuracionTabla = null)
        {
            configuracionTabla = NormalizarConfiguracion(datos, configuracionTabla);

            tituloPagina = (string.IsNullOrEmpty(tituloPagina)) ? "Tabla" : tituloPagina;

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add(tituloPagina);

            IntroducirNotas(ws, configuracionTabla.PosicionInicial, configuracionTabla.Cabeceras);

            IntroducirTabla(ws, datos, configuracionTabla);

            IntroducirNotas(ws, configuracionTabla.PosicionInicial, configuracionTabla.Pies);

            return wb;
        }


        public XLWorkbook CrearDocumentoTheme<T>(List<T> datos, string tituloPagina,
            ConfiguracionTablaTheme configuracionTablaTheme = null)
        {
            var dataTable = datos.ToDataTable();

            return CrearDocumentoTheme(dataTable, tituloPagina, configuracionTablaTheme);
        }

        public XLWorkbook CrearDocumentoTheme(DataTable datos, string tituloPagina,
            ConfiguracionTablaTheme configuracionTablaTheme = null)
        {
            configuracionTablaTheme = NormalizarConfiguracion(datos, configuracionTablaTheme);

            tituloPagina = (string.IsNullOrEmpty(tituloPagina)) ? "Tabla" : tituloPagina;

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add(tituloPagina);

            IntroducirNotas(ws, configuracionTablaTheme.PosicionInicial, configuracionTablaTheme.Cabeceras);

            IntroducirTablaConTheme(ws, datos, configuracionTablaTheme);

            IntroducirNotas(ws, configuracionTablaTheme.PosicionInicial, configuracionTablaTheme.Pies);

            return wb;
        }


        /// <summary>
        /// Introduce los datos de la tabla en la página
        /// usando como configuracion un theme
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="datos"></param>
        /// <param name="configuracionTablaTheme"></param>
        /// <returns></returns>
        private IXLWorksheet IntroducirTablaConTheme(IXLWorksheet worksheet, DataTable datos,
            ConfiguracionTablaTheme configuracionTablaTheme)
        {
            var posicion = SiguientePosicion(worksheet, configuracionTablaTheme.PosicionInicial);

            var tableWithData = worksheet.Cell(posicion.Fila, posicion.Columna).InsertTable(datos.AsEnumerable());
            tableWithData.Theme = configuracionTablaTheme.ThemeTabla;
            tableWithData.SetShowAutoFilter(configuracionTablaTheme.ShowAutoFilter);

            return worksheet;
        }


        /// <summary>
        /// Introduce los datos de la tabla en la página
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="datos"></param>
        /// <param name="configuracionTabla"></param>
        /// <returns></returns>
        private IXLWorksheet IntroducirTabla(IXLWorksheet worksheet, DataTable datos,
            ConfiguracionTabla configuracionTabla)
        {
            var posicion = SiguientePosicion(worksheet, configuracionTabla.PosicionInicial);

            var tableWithData = worksheet.Cell(posicion.Fila, posicion.Columna).InsertTable(datos.AsEnumerable());
            configuracionTabla.EstiloTabla.AplicarEstilo(tableWithData.Style);
            tableWithData.SetShowAutoFilter(configuracionTabla.ShowAutoFilter);

            return worksheet;
        }

        /// <summary>
        /// Introduce una cadena de notas una encima de otra en la página
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="posicionInicial"></param>
        /// <param name="notas"></param>
        /// <returns></returns>
        private IXLWorksheet IntroducirNotas(IXLWorksheet worksheet, PosicionTabla posicionInicial, IList<Nota> notas)
        {
            foreach (var nota in notas)
            {
                var posicion = SiguientePosicion(worksheet, posicionInicial);
                worksheet = IntroducirNota(worksheet, posicion, nota);
            }

            return worksheet;
        }

        /// <summary>
        /// Introduce una nota en la cabecera o pie de la página
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="posicion"></param>
        /// <param name="nota"></param>
        /// <returns></returns>
        private IXLWorksheet IntroducirNota(IXLWorksheet worksheet, PosicionTabla posicion, Nota nota)
        {
            worksheet.Cell(posicion.Fila, posicion.Columna).Value = nota.Texto;

            // Unimos todas las celdas de la nota en una sola.
            var posicionFinal = new PosicionTabla(posicion.Columna + nota.AnchoCeldas - 1, posicion.Fila);

            var rango = worksheet.Range(posicion.Fila, posicion.Columna, posicionFinal.Fila, posicionFinal.Columna)
                .Merge();

            if (nota.Estilo != null)
            {
                nota.Estilo.AplicarEstilo(rango.Style);
            }

            return worksheet;
        }


        /// <summary>
        /// Calcula la siguiente posición libre que habrá
        /// en la tabla excel
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="posicionInicial"></param>
        /// <returns></returns>
        private PosicionTabla SiguientePosicion(IXLWorksheet worksheet, PosicionTabla posicionInicial)
        {
            var celda = worksheet.LastCellUsed();

            if (celda == null)
            {
                return posicionInicial;
            }

            return new PosicionTabla(posicionInicial.Columna, celda.Address.RowNumber + 1);
        }


        /// <summary>
        /// Define el ancho que tendrán las notas de la tabla
        /// </summary>
        /// <param name="datos"></param>
        /// <param name="configuracionTablaTheme"></param>
        /// <returns></returns>
        private void CalcularAnchosNotas(DataTable datos,
            IConfiguracionTabla configuracionTablaTheme)
        {
            var ancho = datos.Columns.Count;

            foreach (var cabecera in configuracionTablaTheme.Cabeceras)
            {
                cabecera.AnchoCeldas = ancho;
            }

            foreach (var nota in configuracionTablaTheme.Pies)
            {
                nota.AnchoCeldas = ancho;
            }
        }

        /// <summary>
        /// Comprueba que los datos del parámetro configuracionTablaTheme
        /// son correctos y les añade a cada pie y cabecera el ancho
        /// que tendrán basándose en el ancho de la propia tabla
        /// </summary>
        /// <param name="datos"></param>
        /// <param name="configuracionTablaTheme"></param>
        /// <returns></returns>
        private ConfiguracionTablaTheme NormalizarConfiguracion(DataTable datos,
            ConfiguracionTablaTheme configuracionTablaTheme)
        {
            configuracionTablaTheme = (configuracionTablaTheme == null) ? new ConfiguracionTablaTheme() : configuracionTablaTheme;

            CalcularAnchosNotas(datos, configuracionTablaTheme);

            return configuracionTablaTheme;
        }


        /// <summary>
        /// Comprueba que los datos del parámetro configuracionTablaTheme
        /// son correctos y les añade a cada pie y cabecera el ancho
        /// que tendrán basándose en el ancho de la propia tabla
        /// </summary>
        /// <param name="datos"></param>
        /// <param name="configuracionTabla"></param>
        /// <returns></returns>
        private ConfiguracionTabla NormalizarConfiguracion(DataTable datos,
            ConfiguracionTabla configuracionTabla)
        {
            configuracionTabla = (configuracionTabla == null) ? new ConfiguracionTabla() : configuracionTabla;

            CalcularAnchosNotas(datos, configuracionTabla);

            return configuracionTabla;
        }

    }
}