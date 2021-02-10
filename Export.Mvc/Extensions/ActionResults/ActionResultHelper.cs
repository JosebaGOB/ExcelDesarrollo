using System;
using System.IO;
using System.Web.Mvc;

namespace Export.Mvc.Extensions.ActionResults
{
    public static class ActionResultHelper
    {
        public static ActionResult ToActionResult(this MemoryStream memoryStream, string nombreFichero)
        {
            return new ExcelResult(nombreFichero, memoryStream);
        }
    }
}