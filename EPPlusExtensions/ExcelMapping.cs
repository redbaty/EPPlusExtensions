using System;
using System.Collections.Generic;
using System.IO;
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
                var excelPropertyMapping = new ExcelPropertyMapping(propertyInfo, null, null, propertyInfo.Name.ToSentence());

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
            var excelPropertyMapping = PropertyMappings.SingleOrDefault(i => i.RuntimeProperty == propertyInfo);
            if (excelPropertyMapping == null)
            {
                excelPropertyMapping = new ExcelPropertyMapping(propertyInfo, null, null, null);
                PropertyMappings.Add(excelPropertyMapping);
            }

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

        public IEnumerable<T> ReadFromExcelFile(Stream fileStream, int sheet = 0, int headerRow = 1)
        {
            using var package = new ExcelPackage();
            package.Load(fileStream);

            using var worksheet = package.Workbook.Worksheets[sheet];
            var maxRows = worksheet.Dimension.Rows;
            var maxColumns = worksheet.Dimension.Columns;
            var columnToPropertyMapping = new Dictionary<int, ExcelPropertyMapping>();
            
            for (var currentRow = headerRow; currentRow <= maxRows; currentRow++)
            {
                var values = Enumerable.Range(1, maxColumns)
                    .Select(i => worksheet.Cells[currentRow, i].Value)
                    .ToArray();
                
                if (currentRow == headerRow)
                {
                    for (var index = 0; index < values.Length; index++)
                    {
                        var headerValue = values[index];
                        var columnHeader = headerValue?.ToString();
                        var excelPropertyMapping = PropertyMappings.SingleOrDefault(i => i.Header == columnHeader);

                        if (excelPropertyMapping != null)
                        {
                            columnToPropertyMapping.Add(index, excelPropertyMapping);
                        }
                    }
                }
                else
                {
                    var newObject = default(T);
                    
                    for (var index = 0; index < values.Length; index++)
                    {
                        if (columnToPropertyMapping.TryGetValue(index, out var mapping))
                        {
                            newObject ??= Activator.CreateInstance<T>();
                            var value = values[index];
                            mapping.RuntimeProperty.SetValue(newObject, mapping.Parse(value, mapping.RuntimeProperty));
                        }
                    }

                    yield return newObject;
                }
            }
        }

        public byte[] WriteExcelFile(IEnumerable<T> items, bool autoFit = true, Action<ExcelRange> headerRowConfig = null)
        {
            using var package = new ExcelPackage();
            using var worksheet = package.Workbook.Worksheets.Add("default");
            WriteToWorksheet(items, worksheet, autoFit, headerRowConfig);
            return package.GetAsByteArray();
        }

        public ExcelWorksheet WriteToWorksheet(IEnumerable<T> items, ExcelWorksheet worksheet, bool autoFit = true, Action<ExcelRange> headerRowConfig = null, int startRow = 1)
        {
            var currentRow = startRow;

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
                worksheet.AutoFit();

            return worksheet;
        }
    }
}