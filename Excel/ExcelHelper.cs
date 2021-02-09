using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;

namespace ConsoleApplication18.Excel
{
    public static class ExcelHelper
    {
        class OrdenColumnasExcel
        {
            public int DataOrder { get; set; }
            public string DataName { get; set; }
            public Type DataType { get; set; }
            public MemberInfo MemberInfo { get; set; }
        }

        /// <summary>
        /// Convierte una lista de objetos
        /// en un datatable.
        /// Extraerá los datos de los campos, métodos y propiedades
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(this IList<T> data)
        {
            var listaColumnas = new List<OrdenColumnasExcel>();
            var members = typeof(T).GetMembers(BindingFlags.Public | BindingFlags.Instance).Where(
                mem => Attribute.IsDefined(mem, typeof(SerializableExcelAttribute))).ToList();

            if (!members.Any())
            {
                throw new Exception("El tipo " + typeof(T).Name + " debe de tener al menos un elemento marcado como 'SerializableExcel'");
            }

            foreach (var memberInfo in members)
            {
                var orden = memberInfo.GetCustomAttribute<SerializableExcelAttribute>();

                listaColumnas.Add(new OrdenColumnasExcel
                {
                    DataName = memberInfo.Name,
                    DataType = GetUnderlyingType(memberInfo),
                    DataOrder = orden.Posicion,
                    MemberInfo = memberInfo
                });
            }

            listaColumnas.Sort((p, q) => p.DataOrder.CompareTo(q.DataOrder));

            DataTable table = new DataTable();

            foreach (var columna in listaColumnas)
            {
                table.Columns.Add(columna.DataName, columna.DataType);
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


        private static Type GetUnderlyingType(this MemberInfo member)
        {
            switch (member.MemberType)
            {
                case MemberTypes.Field:
                    return ((FieldInfo)member).FieldType;
                case MemberTypes.Method:
                    return ((MethodInfo)member).ReturnType;
                case MemberTypes.Property:
                    return ((PropertyInfo)member).PropertyType;
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
                    return ((FieldInfo)memberInfo).GetValue(forObject);
                case MemberTypes.Property:
                    return ((PropertyInfo)memberInfo).GetValue(forObject);
                case MemberTypes.Method:
                    return ((MethodInfo)memberInfo).Invoke(forObject, null);
                default:
                    throw new NotImplementedException();
            }
        }
    }
}