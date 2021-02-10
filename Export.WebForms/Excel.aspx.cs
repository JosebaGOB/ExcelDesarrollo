using System;
using System.Collections.Generic;
using System.IO;
using System.Web;
using ClosedXML.Excel;
using Export.Common.Dto;
using Export.Common.Utils;
using Export.Common.Utils.Excel;

namespace Export.WebForms
{
    public partial class Excel : System.Web.UI.Page
    {
        private IExcelGenerator _excelGenerator;
        private const string NombreFichero = "PruebaExcel";

        protected void Page_Load(object sender, EventArgs e)
        {
            _excelGenerator = new ExcelGenerator();
        }

 
        protected void BtnExcelListaTipada_Click(object sender, EventArgs e)
        {
            IList<PersonaDto> listaPersonas = new List<PersonaDto>();
            PersonaDto persona1 = new PersonaDto("Eduardo", "Goikoa", "Pérez", 34);
            PersonaDto persona2 = new PersonaDto("Amaia", "Goikoa", "Pérez", 45);
            PersonaDto persona3 = new PersonaDto("Arantza", "Asiain", "Muniain", 45);
            listaPersonas.Add(persona1);
            listaPersonas.Add(persona2);
            listaPersonas.Add(persona3);

            var configuracion = new ConfiguracionTablaTheme();
            configuracion.ThemeTabla = XLTableTheme.TableStyleLight9;

            var ficheroExcel = _excelGenerator.CrearMemoryStreamExcel(listaPersonas, NombreFichero, configuracion);
            ImprimirPorPantallaExcel(NombreFichero, ficheroExcel);
        }

        private void ImprimirPorPantallaExcel(string nombreFichero, MemoryStream fichero)
        {
            var httpContext = HttpContext.Current;
            httpContext.Response.ClearContent();
            httpContext.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            httpContext.Response.AddHeader("content-disposition", "attachment;filename=" + nombreFichero + ".xlsx");
            httpContext.Response.ContentEncoding = System.Text.Encoding.GetEncoding("ISO-8859-15");

            using (var stream = fichero)
            {
                stream.WriteTo(httpContext.Response.OutputStream);
                httpContext.Response.End();
            }
        }
    }
}