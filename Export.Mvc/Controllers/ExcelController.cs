using System.Collections.Generic;
using System.Web.Mvc;
using ClosedXML.Excel;
using Export.Common.Dto;
using Export.Common.Utils;
using Export.Common.Utils.Excel;
using Export.Mvc.Extensions.ActionResults;

namespace Export.Mvc.Controllers
{
    public class ExcelController : Controller
    {
        private const string NombreFichero = "PruebaExcel";

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ExcelListaTipada()
        {
            IList<PersonaDto> listaPersonas = new List<PersonaDto>();
            PersonaDto persona1 = new PersonaDto("Eduardo", "Goikoa", "Pérez", 34);
            PersonaDto persona2 = new PersonaDto("Amaia", "Goikoa", "Pérez", 45);
            PersonaDto persona3 = new PersonaDto("Arantza", "Asiain", "Muniain", 45);
            listaPersonas.Add(persona1);
            listaPersonas.Add(persona2);
            listaPersonas.Add(persona3);

            var excelGenerator = new ExcelGenerator();

            var configuracion = new ConfiguracionTablaTheme();
            configuracion.ThemeTabla = XLTableTheme.TableStyleMedium26;

            return excelGenerator.CrearMemoryStreamExcel(listaPersonas, NombreFichero, configuracion).ToActionResult(NombreFichero);
        }
    }
}