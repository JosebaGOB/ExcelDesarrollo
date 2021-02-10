using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;

namespace Export.Common.Utils.Excel
{
    public static class ExcelHelper
    {
        class OrdenColumnasExcel
        {
            public int DataOrder { get; set; }
            public string ColumnName { get; set; }
            public Type ColumnType { get; set; }
            public MemberInfo MemberInfo { get; set; }
        }

        public static MemoryStream ToMemoryStream(this XLWorkbook workbook)
        {
            MemoryStream fs = new MemoryStream();
            workbook.SaveAs(fs);
            fs.Position = 0;
            return fs;
        }


        /// <summary>
        /// Convierte una lista de objetos
        /// en un datatable.
        /// Extraerá los datos de los campos, métodos y propiedades
        /// que están definidos con el atributo SerializableExcel
        /// SI no hay ningún miembro que tenga el atributo,
        /// lanzará una excepción.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(this IEnumerable<T> data)
        {
            var diccionarioColumnas = ObtenerDiccionario<T>(data);

            var listaColumnas = ObtenerColumnasClase(typeof(T), diccionarioColumnas);

            DataTable table = new DataTable();

            foreach (var columna in listaColumnas)
            {
                table.Columns.Add(columna.ColumnName, columna.ColumnType);
            }

            object[] values = new object[listaColumnas.Count];
            foreach (T item in data)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = listaColumnas[i].MemberInfo.GetValue(item);
                }

                table.Rows.Add(values);
            }

            return table;
        }


        private static Dictionary<string, string> ObtenerDiccionario<T>(IEnumerable<T> data)
        {
            var datos = data.ToList();
            if (datos.Any() && typeof(T).GetInterface(typeof(INombresColumnasExcel).FullName) != null)
            {
                return ((INombresColumnasExcel) datos.First()).NombresColumnas;
            }

            return new Dictionary<string, string>();
        }


        /// <summary>
        /// Obtenemos todos los elementos de la clase que contiene la información
        /// que están marcados para serializar en el documento excel.
        /// Usamos miembros porque se permitirá utilizar tanto propiedades
        /// métodos y campos para obtener información que se
        /// añadirá en el excel.
        /// </summary>
        /// <param name="tipo"></param>
        /// <returns></returns>
        private static List<OrdenColumnasExcel> ObtenerColumnasClase(Type tipo, Dictionary<string, string> nombresColumnas)
        {
            var listaColumnas = new List<OrdenColumnasExcel>();

            // Todos los miembros que tengan el atributo de 
            // SerializableExcelAttribute dentro de la clase
            var members = tipo.GetMembers(BindingFlags.Public | BindingFlags.Instance).Where(
                mem => Attribute.IsDefined(mem, typeof(SerializableExcelAttribute))).ToList();

            if (!members.Any())
            {
                throw new Exception("El tipo " + tipo.Name +
                                    " debe de tener al menos un elemento marcado como 'SerializableExcel'");
            }

            // Creamos la lista.
            // Si la clase que estamos usando 
            // implementa INombresColumnasExcel
            // intentaremos buscar el nombre de la
            // columna dentro del diccionario que tiene
            foreach (var memberInfo in members)
            {
                var orden = memberInfo.GetCustomAttribute<SerializableExcelAttribute>();

                string nombreColumna = memberInfo.Name;

                if (nombresColumnas.ContainsKey(nombreColumna))
                {
                    nombreColumna = nombresColumnas[nombreColumna];
                }

                listaColumnas.Add(new OrdenColumnasExcel
                {
                    ColumnName = nombreColumna,
                    ColumnType = GetUnderlyingType(memberInfo),
                    DataOrder = orden.Posicion,
                    MemberInfo = memberInfo
                });
            }

            // Ordenamos las columnas para que estén según el 
            // orden numérico almacenado en SerializableExcel
            listaColumnas.Sort((p, q) => p.DataOrder.CompareTo(q.DataOrder));

            return listaColumnas;
        }


        private static Type GetUnderlyingType(this MemberInfo member)
        {
            switch (member.MemberType)
            {
                case MemberTypes.Field:
                    return ((FieldInfo) member).FieldType;
                case MemberTypes.Method:
                    return ((MethodInfo) member).ReturnType;
                case MemberTypes.Property:
                    return ((PropertyInfo) member).PropertyType;
                default:
                    throw new ArgumentException
                    (
                        "Input MemberInfo must be if type EventInfo, FieldInfo, MethodInfo, or PropertyInfo"
                    );
            }
        }


        private static object GetValue(this MemberInfo memberInfo, object forObject)
        {
            switch (memberInfo.MemberType)
            {
                case MemberTypes.Field:
                    return ((FieldInfo) memberInfo).GetValue(forObject);
                case MemberTypes.Property:
                    return ((PropertyInfo) memberInfo).GetValue(forObject);
                case MemberTypes.Method:
                    return ((MethodInfo) memberInfo).Invoke(forObject, null);
                default:
                    throw new NotImplementedException();
            }
        }
    }
}