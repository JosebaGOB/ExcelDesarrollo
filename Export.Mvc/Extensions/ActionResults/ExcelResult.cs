using System.IO;
using System.Web.Mvc;

namespace Export.Mvc.Extensions.ActionResults
{
    public class ExcelResult : ActionResult
    {
        private string NombreFichero { get; set; }
        private MemoryStream ExcelMemoryStream { get; set; }


        public ExcelResult(string nombreFichero, MemoryStream memoryStream)
        {
            NombreFichero = nombreFichero;
            ExcelMemoryStream = memoryStream;
        }

        public override void ExecuteResult(ControllerContext context)
        {
            if (ExcelMemoryStream == null || ExcelMemoryStream.Length == 0)
                return;

            var response = context.HttpContext.Response;
            response.ContentType =
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            response.AddHeader("content-disposition", "attachment;filename=" + NombreFichero + ".xlsx");

            ExcelMemoryStream.WriteTo(response.OutputStream);
        }
    }
}