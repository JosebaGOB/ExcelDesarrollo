using System;
using System.Collections.Generic;
using Export.Common.Utils.Excel;

namespace Export.Common.Dto
{
    public class PersonaDto : INombresColumnasExcel
    {
        public PersonaDto(string nombre, string apellido1, string apellido2, int edad)
        {
            Nombre = nombre;
            Apellido1 = apellido1;
            Apellido2 = apellido2;
            Edad = edad;
            FechaAlta = DateTime.Now;
        }

        [SerializableExcel(1)] public string Nombre { get; set; }
        [SerializableExcel(2)] public string Apellido1 { get; set; }
        [SerializableExcel(3)] public string Apellido2 { get; set; }
        [SerializableExcel(4)] public int Edad { get; set; }
        [SerializableExcel(5)] public DateTime FechaAlta { get; set; }

        [SerializableExcel(0)]
        public string NombreCompleto()
        {
            return Nombre + " " + Apellido1 + " " + Apellido2;
        }

        public Dictionary<string, string> NombresColumnas
        {
            get
            {
                return new Dictionary<string, string>
                {
                    {"Nombre", "Nombre"},
                    {"Apellido1", "Primer apellido"},
                    {"Apellido2", "Segundo apellido"},
                    {"NombreCompleto", "Nombre completo"}
                };
            }
        }
    }
}