using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using OfficeOpenXml;

namespace EPPlusExtensions
{
    public class ExcelMappingWithItems<T> : ExcelMapping<T>
    {
        public ExcelMappingWithItems(IEnumerable<T> items)
        {
            Items = items;
        }

        private IEnumerable<T> Items { get; }

        public byte[] WriteExcelFile(bool autoFit = true, Action<ExcelRange> headerRowConfig = null)
        {
            return WriteExcelFile(Items, autoFit, headerRowConfig);
        }

        public new ExcelMappingWithItems<T> AutoMap()
        {
            return (ExcelMappingWithItems<T>) base.AutoMap();
        }

        public new ExcelMappingWithItems<T> Property<TObj>(Expression<Func<T, TObj>> propertyLambda,
                                                           Action<ExcelPropertyMapping> action)
        {
            return (ExcelMappingWithItems<T>) base.Property(propertyLambda, action);
        }

        public new ExcelMappingWithItems<T> RemovePropertyMapping<TObj>(Expression<Func<T, TObj>> propertyLambda)
        {
            return (ExcelMappingWithItems<T>) base.RemovePropertyMapping(propertyLambda);
        }
    }
}