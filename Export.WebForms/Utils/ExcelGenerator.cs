using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace Export.WebForms.Utils
{
    //public class ExcelGenerator : IExcelGenerator
    //{
    //    public MemoryStream CreateExcelFile(string name, IEnumerable<object> data)
    //    {
    //        var workbook = new XLWorkbook();
    //        var worksheet = workbook.AddWorksheet(name);
    //        var row = 1;
    //        var col = 1;

    //        var type = data.First().GetType();
    //        var props = type.GetProperties();

    //        foreach (var prop in props)
    //        {
    //            var cell = worksheet.Cell(row, col++);
    //            cell.Value = prop.Name;
    //            cell.Style.Font.Bold = true;
    //            cell.Style.Font.FontColor = XLColor.Blue;
    //        }

    //        foreach (var elem in data)
    //        {
    //            row++;
    //            col = 1;
    //            foreach (var prop in props)
    //            {
    //                var cell = worksheet.Cell(row, col++);
    //                cell.Value = prop.GetValue(elem).ToString();
    //            }
    //        }

    //        MemoryStream memoryStream = new MemoryStream();
    //        workbook.SaveAs(memoryStream);

    //        return memoryStream;
    //    }
    //}
}