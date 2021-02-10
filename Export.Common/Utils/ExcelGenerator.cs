using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Export.Common.Utils.Excel;

namespace Export.Common.Utils
{
    public class ExcelGenerator : IExcelGenerator
    {
        /// <summary>
        /// Crea un archivo Excel en formato de array de bytes.
        /// </summary>
        /// <param name="datos">datos para crear el archivo</param>
        /// <param name="tituloPagina">Nombre de la página que contendrá la tabla</param>
        /// <param name="configuracionTabla">Configuración del diseño de la página</param>
        /// <returns>Array de bytes con el contenido del archivo ya creado</returns>
        public MemoryStream CrearMemoryStreamExcel<T>(IEnumerable<T> datos, string tituloPagina,
            IConfiguracionTabla configuracionTabla = null)
        {
            return CrearMemoryStreamExcel(datos.ToDataTable(), tituloPagina, configuracionTabla);
        }

        /// <summary>
        /// Crea un archivo Excel en formato de array de bytes.
        /// </summary>
        /// <param name="datos">datos para crear el archivo</param>
        /// <param name="tituloPagina">Nombre de la página que contendrá la tabla</param>
        /// <param name="configuracionTabla">Configuración del diseño de la página</param>
        /// <returns>Array de bytes con el contenido del archivo ya creado</returns>
        public MemoryStream CrearMemoryStreamExcel(DataTable datos, string tituloPagina,
            IConfiguracionTabla configuracionTabla = null)
        {
            XLWorkbook documento;

            if (configuracionTabla == null || configuracionTabla.GetType() == typeof(ConfiguracionTablaManual))
            {
                documento = CrearDocumento(datos, tituloPagina, (ConfiguracionTablaManual) configuracionTabla);
            }
            else if (configuracionTabla.GetType() == typeof(ConfiguracionTablaTheme))
            {
                documento = CrearDocumento(datos, tituloPagina, (ConfiguracionTablaTheme) configuracionTabla);
            }
            else
            {
                throw new Exception("El tipo " + configuracionTabla.GetType().Name +
                                    " como configuración de tablas Excel no está soportado.");
            }

            return documento.ToMemoryStream();
        }

        /// <summary>
        /// Crea una instancia de un documento Excel usando una instancia de configuración
        /// para su diseño.
        /// </summary>
        /// <param name="datos">datos para crear el archivo</param>
        /// <param name="tituloPagina">Nombre de la página que contendrá la tabla</param>
        /// <param name="configuracionTablaManual">Configuración del diseño de la página</param>
        /// <returns>Array de bytes con el contenido del archivo ya creado</returns>
        public XLWorkbook CrearDocumento<T>(IEnumerable<T> datos, string tituloPagina,
            ConfiguracionTablaManual configuracionTablaManual = null)
        {
            var dataTable = datos.ToDataTable();

            return CrearDocumento(dataTable, tituloPagina, configuracionTablaManual);
        }


        /// <summary>
        /// Crea una instancia de un documento Excel usando una instancia de configuración
        /// para su diseño.
        /// </summary>
        /// <param name="datos">datos para crear el archivo</param>
        /// <param name="tituloPagina">Nombre de la página que contendrá la tabla</param>
        /// <param name="configuracionTablaManual">Configuración del diseño de la página</param>
        /// <returns>Array de bytes con el contenido del archivo ya creado</returns>
        public XLWorkbook CrearDocumento(DataTable datos, string tituloPagina,
            ConfiguracionTablaManual configuracionTablaManual = null)
        {
            configuracionTablaManual = NormalizarConfiguracion(datos, configuracionTablaManual);

            tituloPagina = (string.IsNullOrEmpty(tituloPagina)) ? "Tabla" : tituloPagina;

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add(tituloPagina);

            IntroducirNotas(ws, configuracionTablaManual.PosicionInicial, configuracionTablaManual.Cabeceras);

            IntroducirTabla(ws, datos, configuracionTablaManual);

            IntroducirNotas(ws, configuracionTablaManual.PosicionInicial, configuracionTablaManual.Pies);

            return wb;
        }


        /// <summary>
        /// Crea una instancia de un documento Excel usando una instancia de configuración
        /// para su diseño.
        /// </summary>
        /// <param name="datos">datos para crear el archivo</param>
        /// <param name="tituloPagina">Nombre de la página que contendrá la tabla</param>
        /// <param name="configuracionTablaTheme">Configuración del diseño de la página</param>
        /// <returns>Array de bytes con el contenido del archivo ya creado</returns>
        public XLWorkbook CrearDocumento<T>(IEnumerable<T> datos, string tituloPagina,
            ConfiguracionTablaTheme configuracionTablaTheme = null)
        {
            var dataTable = datos.ToDataTable();

            return CrearDocumento(dataTable, tituloPagina, configuracionTablaTheme);
        }

        /// <summary>
        /// Crea una instancia de un documento Excel usando una instancia de configuración
        /// para su diseño.
        /// </summary>
        /// <param name="datos">datos para crear el archivo</param>
        /// <param name="tituloPagina">Nombre de la página que contendrá la tabla</param>
        /// <param name="configuracionTablaTheme">Configuración del diseño de la página</param>
        /// <returns>Array de bytes con el contenido del archivo ya creado</returns>
        public XLWorkbook CrearDocumento(DataTable datos, string tituloPagina,
            ConfiguracionTablaTheme configuracionTablaTheme = null)
        {
            configuracionTablaTheme = NormalizarConfiguracion(datos, configuracionTablaTheme);

            tituloPagina = (string.IsNullOrEmpty(tituloPagina)) ? "Tabla" : tituloPagina;

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add(tituloPagina);

            IntroducirNotas(ws, configuracionTablaTheme.PosicionInicial, configuracionTablaTheme.Cabeceras);

            IntroducirTabla(ws, datos, configuracionTablaTheme);

            IntroducirNotas(ws, configuracionTablaTheme.PosicionInicial, configuracionTablaTheme.Pies);

            return wb;
        }

        /// <summary>
        /// Introduce los datos de la tabla en la página
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="datos"></param>
        /// <param name="configuracionTabla"></param>
        /// <returns></returns>
        private IXLWorksheet IntroducirTabla(IXLWorksheet worksheet, DataTable datos,
            IConfiguracionTabla configuracionTabla)
        {
            var posicion = SiguientePosicion(worksheet, configuracionTabla.PosicionInicial);

            var tableWithData = worksheet.Cell(posicion.Fila, posicion.Columna).InsertTable(datos.AsEnumerable());

            if (configuracionTabla.GetType() == typeof(ConfiguracionTablaManual))
            {
                tableWithData.Theme = XLTableTheme.None;
                ((ConfiguracionTablaManual) configuracionTabla).EstiloTabla.AplicarEstilo(tableWithData.Style);
            }
            else if (configuracionTabla.GetType() == typeof(ConfiguracionTablaTheme))
            {
                tableWithData.Theme = ((ConfiguracionTablaTheme) configuracionTabla).ThemeTabla;
            }

            tableWithData.SetShowAutoFilter(configuracionTabla.ShowAutoFilter);
            worksheet.Columns().AdjustToContents();
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
            configuracionTablaTheme = (configuracionTablaTheme == null)
                ? new ConfiguracionTablaTheme()
                : configuracionTablaTheme;

            CalcularAnchosNotas(datos, configuracionTablaTheme);

            return configuracionTablaTheme;
        }


        /// <summary>
        /// Comprueba que los datos del parámetro configuracionTablaTheme
        /// son correctos y les añade a cada pie y cabecera el ancho
        /// que tendrán basándose en el ancho de la propia tabla
        /// </summary>
        /// <param name="datos"></param>
        /// <param name="configuracionTablaManual"></param>
        /// <returns></returns>
        private ConfiguracionTablaManual NormalizarConfiguracion(DataTable datos,
            ConfiguracionTablaManual configuracionTablaManual)
        {
            configuracionTablaManual = (configuracionTablaManual == null)
                ? new ConfiguracionTablaManual()
                : configuracionTablaManual;

            CalcularAnchosNotas(datos, configuracionTablaManual);

            return configuracionTablaManual;
        }
    }
}