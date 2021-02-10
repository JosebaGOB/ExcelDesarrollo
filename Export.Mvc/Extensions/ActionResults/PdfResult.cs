using System;
using System.IO;
using System.Web.Mvc;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Export.Mvc.Extensions.ActionResults
{
    public class PdfFileResult : ActionResult
    {
        private readonly string _fileName;
        private readonly string _text;

        public PdfFileResult(string fileName, string text)
        {
            _fileName = fileName;
            if (string.IsNullOrWhiteSpace(_fileName))
                _fileName = "document.pdf";
            else if (!_fileName.ToLower().EndsWith(".pdf"))
                _fileName += ".pdf";

            _text = string.IsNullOrWhiteSpace(text) ? "Default text" : text;
        }

        public override void ExecuteResult(ControllerContext context)
        {
            context.HttpContext.Response.ContentType = "application/pdf";
            context.HttpContext.Response.AddHeader("content-disposition", "attachment; filename=" + _fileName);
            generatePdf(context.HttpContext.Response.OutputStream, _text);
        }

        private void generatePdf(Stream stream, string text)
        {
            var document = new Document();
            PdfWriter writer = PdfWriter.GetInstance(document, stream);
            document.Open();

            document.Add(new Paragraph(text, FontFactory.GetFont("Arial", 28, BaseColor.GRAY)));
            document.Add(new Paragraph(DateTime.Now.ToLongDateString()));
            document.Close();
            writer.Flush();
            writer.Close();
        }

    }
}