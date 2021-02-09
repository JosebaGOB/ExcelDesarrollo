using System;

namespace ConsoleApplication18.Excel
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Method | AttributeTargets.Field)]
    public class SerializableExcelAttribute : Attribute
    {
        private int _posicion;

        public SerializableExcelAttribute(int posicion = 0)
        {
            _posicion = posicion;
        }

        public int Posicion
        {
            get { return _posicion; }
        }
    }
}