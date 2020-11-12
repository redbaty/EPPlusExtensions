using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlusExtensions
{
    public class ManualExcelHeader
    {
        public ManualExcelHeader(string header,
            int column)
        {
            Header = header;
            Column = column;
        }

        public int Column { get; }

        public string Header { get; }

        public override string ToString() => $"{Header} [{Column}]";
    }
    
    public class WorksheetWithManualHeaders : IDisposable
    {
        public WorksheetWithManualHeaders(ExcelWorksheet sheet) : this(sheet, Enumerable.Empty<ManualExcelHeader>())
        {
        }

        public WorksheetWithManualHeaders(ExcelWorksheet sheet, IEnumerable<ManualExcelHeader> headers)
        {
            Sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            Headers = headers.ToDictionary(i => i.Header);
        }

        public ExcelWorksheet Sheet { get; }

        internal Dictionary<string, ManualExcelHeader> Headers { get; }

        public void Dispose()
        {
            Sheet?.Dispose();
        }
    }
    
    public static class ManualExcelExtensions
    {
        public static WorksheetWithManualHeaders AddHeader(this WorksheetWithManualHeaders sheet,
            params string[] headers)
        {
            foreach (var header in headers)
                sheet.Headers.Add(header,
                    new ManualExcelHeader(TreatHeader(header),
                        GetHighestColumn(sheet.Headers.Values) + 1));

            return sheet;
        }

        public static void AutoFit(this IEnumerable<ExcelWorksheet> sheets)
        {
            foreach (var excelWorksheet in sheets) excelWorksheet.AutoFit();
        }

        public static void AutoFit(this WorksheetWithManualHeaders manualHeaders)
        {
            manualHeaders.Sheet.AutoFit();
        }

        public static void AutoFit(this ExcelWorksheet sheet)
        {
            if (sheet.Dimension != null)
                sheet.Cells[sheet.Dimension.Address]
                    .AutoFitColumns();
        }

        public static int GetLastRow(this WorksheetWithManualHeaders manualHeaders) => GetLastRow(manualHeaders.Sheet);

        public static int GetLastRow(this ExcelWorksheet worksheet) => worksheet.Dimension?.Rows ?? 0;

        public static bool HasHeader(this WorksheetWithManualHeaders sheet,
            params string[] headers)
        {
            return headers.All(i => sheet.Headers.ContainsKey(i));
        }

        public static WorksheetWithManualHeaders LoadHeaders(this WorksheetWithManualHeaders sheet,
            int headersRow)
        {
            foreach (var i in Enumerable.Range(1,
                sheet.Sheet.Dimension.Columns))
            {
                var cell = sheet.Sheet.Cells[headersRow,
                    i];
                var headerValue = cell.Value?.ToString();

                if (!string.IsNullOrEmpty(headerValue)) sheet.AddHeader(headerValue);
            }

            return sheet;
        }

        public static void SetColor(this ExcelWorksheet sheet,
            int row,
            Color color,
            ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            var cell = sheet.Row(row)
                .Style.Fill;
            cell.PatternType = fillStyle;
            cell.BackgroundColor.SetColor(color);
        }

        public static WorksheetWithManualHeaders WithHeaders(this ExcelWorksheet sheet) => new WorksheetWithManualHeaders(sheet);

        public static WorksheetWithManualHeaders WriteHeaders(this WorksheetWithManualHeaders sheet,
            int row = 1)
        {
            foreach (var excelHeader in sheet.Headers.Values)
            {
                var cell = sheet.Sheet.Cells[row,
                    excelHeader.Column];
                cell.Value = excelHeader.Header;
                cell.Style.Font.Bold = true;
            }

            return sheet;
        }

        public static void WriteToColumn(this WorksheetWithManualHeaders sheet,
            int row,
            string columnHeader,
            object value,
            string format = null)
        {
            if (!sheet.Headers.ContainsKey(columnHeader)) return;

            var header = sheet.Headers[columnHeader];

            WriteToColumn(sheet,
                row,
                header.Column,
                value,
                format);
        }

        public static ExcelRange WriteToColumn(this WorksheetWithManualHeaders sheet,
            int row,
            int column,
            object value,
            string format = null)
        {
            var cell = sheet.Sheet.Cells[row, column];

            if (string.IsNullOrEmpty(format))
            {
                if (value is DateTime) cell.Style.Numberformat.Format = "dd/mm/yyyy";

                if (value is decimal || value is double) cell.AddDecimalFormat();
            }
            else
            {
                cell.Style.Numberformat.Format = format;
            }

            cell.Value = value;
            return cell;
        }

        public static IEnumerable<ExcelWorksheet> WriteToSheet<T>(this ExcelWorkbook workbook,
            IEnumerable<T> items,
            Func<Type, string> keyGenerator)
        {
            var sheetDictionary = new Dictionary<Type, ExcelWorksheet>();

            foreach (var item in items)
            {
                var itemType = item.GetType();
                var properties = itemType.GetProperties()
                    .Where(i => i.CanRead)
                    .ToList();

                if (!sheetDictionary.ContainsKey(itemType))
                    sheetDictionary.Add(itemType,
                        workbook.GetOrCreate(keyGenerator.Invoke(itemType),
                            worksheet => WriteHeader(properties,
                                worksheet)));

                var sheet = sheetDictionary[itemType];
                var lastRow = sheet.GetLastRow() + 1;

                for (var index = 0; index < properties.Count; index++)
                {
                    var propertyInfo = properties[index];
                    
                    var cell = sheet.Cells[lastRow,
                        index + 1];
                    var value = propertyInfo.GetValue(item);

                    if (value is DateTime) cell.Style.Numberformat.Format = "dd/mm/yyyy";

                    cell.Value = value;
                }
            }

            return sheetDictionary.Values;
        }

        private static void AddDecimalFormat(this ExcelRange range)
        {
            range.Style.Numberformat.Format = "#,##0.00";
        }

        private static ExcelWorksheet CreateSheet(ExcelWorkbook workbook,
            string key,
            Action<ExcelWorksheet> onCreate)
        {
            var newSheet = workbook.Worksheets.Add(key);
            onCreate?.Invoke(newSheet);
            return newSheet;
        }

        private static int GetHighestColumn(IReadOnlyCollection<ManualExcelHeader> headers)
        {
            return !headers.Any() ? 0 : headers.Max(i => i.Column);
        }

        private static ExcelWorksheet GetOrCreate(this ExcelWorkbook workbook,
            string key,
            Action<ExcelWorksheet> onCreate)
        {
            return workbook.Worksheets.FirstOrDefault(i => string.Equals(i.Name,
                key,
                StringComparison.InvariantCultureIgnoreCase)) ?? CreateSheet(workbook,
                key,
                onCreate);
        }

        private static string TreatHeader(string header)
        {
            header = header.Trim();
            header = header.Replace("\n",
                " ");
            header = header.Replace("\r",
                string.Empty);
            header = Regex.Replace(header,
                @"\s+",
                " ");
            return header;
        }

        private static void WriteHeader(IReadOnlyList<PropertyInfo> properties,
            ExcelWorksheet worksheet)
        {
            for (var index = 0;
                index < properties.Count;
                index++)
            {
                var propertyInfo = properties[index];
                var cell = worksheet.Cells[1,
                    index + 1];
                cell.Value = propertyInfo.Name;
                cell.Style.Font.Bold = true;
            }
        }
    }
}