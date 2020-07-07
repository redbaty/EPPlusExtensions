using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using EPPlusExtensions.Annotations;
using EPPlusExtensions.Extensions;
using OfficeOpenXml;

namespace EPPlusExtensions
{
    public class ExcelMapping<T>
    {
        public HashSet<ExcelPropertyMapping> PropertyMappings { get; } = new HashSet<ExcelPropertyMapping>();

        public ExcelMapping<T> AutoMap()
        {
            foreach (var propertyInfo in typeof(T).GetProperties())
            {
                var excelPropertyMapping = new ExcelPropertyMapping(propertyInfo, null, propertyInfo.Name);

                if (propertyInfo.GetCustomAttribute<ExcelColumnAttribute>() is {} excelColumnAttribute)
                {
                    excelPropertyMapping.Header = excelColumnAttribute.Column;
                }

                if (propertyInfo.GetCustomAttribute<ExcelFormatAttribute>() is {} excelFormatAttribute)
                {
                    excelPropertyMapping.Header = excelFormatAttribute.Format;
                }

                if (propertyInfo.GetCustomAttribute<ExcelOrderAttribute>() is {} excelOrderAttribute)
                {
                    excelPropertyMapping.Order = excelOrderAttribute.Order;
                }

                PropertyMappings.Add(excelPropertyMapping);
            }

            return this;
        }

        public ExcelMapping<T> Property<TObj>(Expression<Func<T, TObj>> propertyLambda,
                                              Action<ExcelPropertyMapping> action)
        {
            var propertyInfo = propertyLambda.GetProperty();
            var excelPropertyMapping = PropertyMappings.SingleOrDefault(i => i.RuntimeProperty == propertyInfo) ??
                                       new ExcelPropertyMapping(propertyInfo, null, null);
            action.Invoke(excelPropertyMapping);
            return this;
        }

        public ExcelMapping<T> RemovePropertyMapping<TObj>(Expression<Func<T, TObj>> propertyLambda)
        {
            var propertyInfo = propertyLambda.GetProperty();
            var propertyMapping = PropertyMappings.SingleOrDefault(i => i.RuntimeProperty == propertyInfo);

            if (propertyMapping == null)
            {
                return this;
            }

            PropertyMappings.Remove(propertyMapping);
            return this;
        }

        public byte[] WriteExcelFile(IEnumerable<T> items, bool autoFit = true, Action<ExcelRange> headerRowConfig = null)
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("default");
            var currentRow = 1;

            var sortedHeaders = PropertyMappings.OrderByDescending(i => i.Order).Select(i => i.Header)
                .Distinct()
                .Select((excel, index) => new {Index = index, Value = excel})
                .ToDictionary(i => i.Index + 1, i => i.Value);

            var mapHeaders = sortedHeaders.ToDictionary(i => i.Value, i => i.Key);

            foreach (var (key, value) in sortedHeaders)
            {
                var worksheetCell = worksheet.Cells[currentRow, key];
                worksheetCell.Value = value;
                headerRowConfig?.Invoke(worksheetCell);
            }

            currentRow++;

            foreach (var item in items)
            {
                foreach (var propertyMapping in PropertyMappings)
                {
                    var value = propertyMapping.GetValue(item);
                    worksheet.Cells[currentRow, mapHeaders[propertyMapping.Header]].Value = value;

                    if (!string.IsNullOrEmpty(propertyMapping.Format))
                        worksheet.Cells[currentRow, mapHeaders[propertyMapping.Header]].Style.Numberformat.Format =
                            propertyMapping.Format;
                }

                currentRow++;
            }

            if (autoFit)
                package.AutoFit();

            return package.GetAsByteArray();
        }
    }
}