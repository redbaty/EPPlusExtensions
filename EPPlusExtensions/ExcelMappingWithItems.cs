using System.Collections.Generic;

namespace EPPlusExtensions
{
    public class ExcelMappingWithItems<T> : ExcelMapping<T>
    {
        public ExcelMappingWithItems(IEnumerable<T> items)
        {
            Items = items;
        }

        private IEnumerable<T> Items { get; }

        public byte[] WriteExcelFile(bool autoFit = true)
        {
            return WriteExcelFile(Items, autoFit);
        }
    }
}