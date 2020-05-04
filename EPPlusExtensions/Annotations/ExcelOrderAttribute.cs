using System;

namespace EPPlusExtensions.Annotations
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelOrderAttribute : Attribute
    {
        public ExcelOrderAttribute(int order)
        {
            Order = order;
        }

        internal int Order { get; }
    }
}