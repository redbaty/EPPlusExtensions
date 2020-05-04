using System;

namespace EPPlusExtensions.Annotations
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        public ExcelColumnAttribute(string column)
        {
            Column = column;
        }

        internal string Column { get; }
    }
}