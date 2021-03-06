﻿using System.Collections.Generic;
using OfficeOpenXml;

namespace EPPlusExtensions.Extensions
{
    public static class ExcelExtensions
    {
        public static void AutoFit(this ExcelPackage package)
        {
            package.Workbook.AutoFit();
        }

        public static void AutoFit(this ExcelWorkbook workbook)
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                worksheet.AutoFit();
            }
        }

        public static void AutoFit(this ExcelWorksheet sheet)
        {
            if (sheet.Dimension != null)
                sheet.Cells[sheet.Dimension.Address]
                    .AutoFitColumns();
        }
        
        public static ExcelMappingWithItems<T> CreateMapping<T>(this IEnumerable<T> items)
        {
            return new ExcelMappingWithItems<T>(items);
        }
    }
}