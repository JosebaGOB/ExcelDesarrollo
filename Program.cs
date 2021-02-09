using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ConsoleApplication18.Excel;

namespace ConsoleApplication18
{
    class Program
    {
        static void Main(string[] args)
        {
            // var a = new Fuente("Calibri", 9);

            var derp = new TablaExcel();

            var configuracionTablaTheme = new ConfiguracionTablaTheme();

            configuracionTablaTheme.Cabeceras = new List<Nota>
            {
                new Nota("hola")
            };

            var estiloPie = new Estilo();
            estiloPie.Bordes = new Bordes
            {
                Color = XLColor.AmberSaeEce,
                Posicion = PosicionBordes.Todos
            };
            estiloPie.Relleno = new Relleno
            {
                ColorFondo = XLColor.Yellow
            };


            configuracionTablaTheme.Pies = new List<Nota>
            {
                new Nota("adios", estiloPie)
            };
            configuracionTablaTheme.ThemeTabla = XLTableTheme.TableStyleLight15;
            configuracionTablaTheme.ShowAutoFilter = true;


            var list = new List<Person>();
            list.Add(new Person() { Name = "John", Age = 30, House = "On Elm St." });
            list.Add(new Person() { Name = "Mary", Age = 15, House = "On Main St." });
            list.Add(new Person() { Name = "Luis", Age = 21, House = "On 23rd St." });
            list.Add(new Person() { Name = "Henry", Age = 45, House = "On 5th Ave." });



            var configuracionTabla = new ConfiguracionTabla();

            configuracionTabla.EstiloTabla = new Estilo
            {
                Fuente = new Fuente
                {
                    Color = XLColor.Red
                },
                Bordes = new Bordes
                {
                    Estilo = XLBorderStyleValues.Double,
                    Color = XLColor.Blue,
                    Posicion = PosicionBordes.Todos
                }
            };

            configuracionTabla.Cabeceras = new List<Nota>
            {
                new Nota("hola")
            };

            estiloPie = new Estilo();
            estiloPie.Bordes = new Bordes
            {
                Color = XLColor.AmberSaeEce,
                Posicion = PosicionBordes.Todos
            };
            estiloPie.Relleno = new Relleno
            {
                ColorFondo = XLColor.Yellow
            };


            configuracionTabla.Pies = new List<Nota>
            {
                new Nota("adios", estiloPie)
            };
            configuracionTabla.ShowAutoFilter = true;

            var excelTheme = derp.CrearDocumentoTheme(list, "", configuracionTablaTheme);
            excelTheme.SaveAs("excelTheme.xlsx");

            var excel = derp.CrearDocumento(list, "", configuracionTabla);
            excel.SaveAs("excelConfigurado.xlsx");
        }


        public static void Create()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Inserting Tables");

            // From a list of strings
            var listOfStrings = new List<String>();
            listOfStrings.Add("House");
            listOfStrings.Add("Car");
            ws.Cell(1, 1).Value = "From Strings";
            ws.Cell(1, 1).AsRange().AddToNamed("Titles");
            var tableWithStrings = ws.Cell(2, 1).InsertTable(listOfStrings);

            // From a list of arrays
            var listOfArr = new List<Int32[]>();
            listOfArr.Add(new Int32[] {1, 2, 3});
            listOfArr.Add(new Int32[] {1});
            listOfArr.Add(new Int32[] {1, 2, 3, 4, 5, 6});
            ws.Cell(1, 3).Value = "From Arrays";
            ws.Range(1, 3, 1, 8).Merge().AddToNamed("Titles");
            var tableWithArrays = ws.Cell(2, 3).InsertTable(listOfArr);
            tableWithArrays.Theme = XLTableTheme.TableStyleDark1;
            // From a DataTable
            var dataTable = GetTable();
            ws.Cell(7, 1).Value = "From DataTable";
            ws.Range(7, 1, 7, 4).Merge().AddToNamed("Titles");
            var tableWithData = ws.Cell(8, 1).InsertTable(dataTable.AsEnumerable());
            tableWithData.ShowRowStripes = true;
            tableWithData.Theme = XLTableTheme.None;

            // From a query
            var list = new List<Person>();
            list.Add(new Person() {Name = "John", Age = 30, House = "On Elm St."});
            list.Add(new Person() {Name = "Mary", Age = 15, House = "On Main St."});
            list.Add(new Person() {Name = "Luis", Age = 21, House = "On 23rd St."});
            list.Add(new Person() {Name = "Henry", Age = 45, House = "On 5th Ave."});

            var people = from p in list
                where p.Age >= 21
                select new {p.Name, p.House, p.Age};
            ws.Cell(7, 6).Value = "From Query";
            ws.Range(7, 6, 7, 8).Merge().AddToNamed("Titles");
            var tableWithPeople = ws.Cell(8, 6).InsertTable(people.AsEnumerable());
            //  tableWithPeople.Style.Fill.BackgroundColor = XLColor.AliceBlue;
            // Prepare the style for the titles
            var titlesStyle = wb.Style;
            titlesStyle.Font.Bold = true;
            titlesStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titlesStyle.Fill.BackgroundColor = XLColor.Cyan;

            // Format all titles in one shot
            wb.NamedRanges.NamedRange("Titles").Ranges.Style = titlesStyle;

            ws.Columns().AdjustToContents();

            wb.SaveAs("InsertingTables.xlsx");
            Stream fs = new MemoryStream();
            wb.SaveAs(fs);
            fs.Position = 0;
        }


        class Person
        {
            [SerializableExcel()]
            public String House { get; set; }
            [SerializableExcel()]
            public String Name { get; set; }
            [SerializableExcel()]
            public Int32 Age { get; set; }
        }

        private static DataTable GetTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            table.Rows.Add(25, "Indocin", "David", DateTime.Now);
            table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now);
            table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now);
            table.Rows.Add(21, "Combivent", "Janet", DateTime.Now);
            table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now);
            return table;
        }
    }
}