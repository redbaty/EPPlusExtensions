using System;

namespace EPPlusExtensions.Annotations
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelFormatAttribute : Attribute
    {
        public ExcelFormatAttribute(string format)
        {
            Format = format;
        }

        internal string Format { get; }
    }
}