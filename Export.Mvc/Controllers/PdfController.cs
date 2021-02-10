using System.Web.Mvc;
using Export.Common.Utils;
using Export.Mvc.Extensions.ActionResults;

namespace Export.Mvc.Controllers
{
    public class PdfController : Controller
    {

        public ActionResult Index()
        {
            return View();
        }

        //
        // POST: /Home/CrearPdf
        [HttpPost]
        public ActionResult CrearPdf(string nombreArchivo, string texto)
        {
            return new PdfFileResult(nombreArchivo, texto);
        }
    }
}