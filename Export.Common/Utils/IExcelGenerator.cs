using System.Collections.Generic;
using System.Data;
using System.IO;
using ClosedXML.Excel;
using Export.Common.Utils.Excel;

namespace Export.Common.Utils
{
    public interface IExcelGenerator
    {
        /// <summary>
        /// Crea un archivo Excel en formato de array de bytes.
        /// </summary>
        /// <param name="datos">datos para crear el archivo</param>
        /// <param name="tituloPagina">Nombre de la página que contendrá la tabla</param>
        /// <param name="configuracionTabla">Configuración del diseño de la página</param>
        /// <returns>Array de bytes con el contenido del archivo ya creado</returns>
        MemoryStream CrearMemoryStreamExcel<T>(IEnumerable<T> datos, string tituloPagina,
            IConfiguracionTabla configuracionTabla = null);

        /// <summary>
        /// Crea un archivo Excel en formato de array de bytes.
        /// </summary>
        /// <param name="datos">datos para crear el archivo</param>
        /// <param name="tituloPagina">Nombre de la página que contendrá la tabla</param>
        /// <param name="configuracionTabla">Configuración del diseño de la página</param>
        /// <returns>Array de bytes con el contenido del archivo ya creado</returns>
        MemoryStream CrearMemoryStreamExcel(DataTable datos, string tituloPagina,
            IConfiguracionTabla configuracionTabla = null);

        /// <summary>
        /// Crea una instancia de un documento Excel usando una instancia de configuración
        /// para su diseño.
        /// </summary>
        /// <param name="datos">datos para crear el archivo</param>
        /// <param name="tituloPagina">Nombre de la página que contendrá la tabla</param>
        /// <param name="configuracionTablaManual">Configuración del diseño de la página</param>
        /// <returns>Array de bytes con el contenido del archivo ya creado</returns>
        XLWorkbook CrearDocumento<T>(IEnumerable<T> datos, string tituloPagina,
            ConfiguracionTablaManual configuracionTablaManual = null);

        /// <summary>
        /// Crea una instancia de un documento Excel usando una instancia de configuración
        /// para su diseño.
        /// </summary>
        /// <param name="datos">datos para crear el archivo</param>
        /// <param name="tituloPagina">Nombre de la página que contendrá la tabla</param>
        /// <param name="configuracionTablaManual">Configuración del diseño de la página</param>
        /// <returns>Array de bytes con el contenido del archivo ya creado</returns>
        XLWorkbook CrearDocumento(DataTable datos, string tituloPagina,
            ConfiguracionTablaManual configuracionTablaManual = null);

        /// <summary>
        /// Crea una instancia de un documento Excel usando una instancia de configuración
        /// para su diseño.
        /// </summary>
        /// <param name="datos">datos para crear el archivo</param>
        /// <param name="tituloPagina">Nombre de la página que contendrá la tabla</param>
        /// <param name="configuracionTablaTheme">Configuración del diseño de la página</param>
        /// <returns>Array de bytes con el contenido del archivo ya creado</returns>
        XLWorkbook CrearDocumento<T>(IEnumerable<T> datos, string tituloPagina,
            ConfiguracionTablaTheme configuracionTablaTheme = null);

        /// <summary>
        /// Crea una instancia de un documento Excel usando una instancia de configuración
        /// para su diseño.
        /// </summary>
        /// <param name="datos">datos para crear el archivo</param>
        /// <param name="tituloPagina">Nombre de la página que contendrá la tabla</param>
        /// <param name="configuracionTablaTheme">Configuración del diseño de la página</param>
        /// <returns>Array de bytes con el contenido del archivo ya creado</returns>
        XLWorkbook CrearDocumento(DataTable datos, string tituloPagina,
            ConfiguracionTablaTheme configuracionTablaTheme = null);
    }
}