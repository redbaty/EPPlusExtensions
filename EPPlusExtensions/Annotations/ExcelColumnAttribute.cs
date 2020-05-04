using System;

namespace EPPlusExtensions.Annotations
{
    public class ExcelColumnAttribute : Attribute
    {
        public ExcelColumnAttribute(string column)
        {
            Column = column;
        }

        internal string Column { get; }
    }
}