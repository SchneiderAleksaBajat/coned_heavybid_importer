using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;


namespace ConEd.HeavyBid.Importer.Utility
{
    public static class ExcelHelper
    {
        public static string GetCellValue(Excel.Range colRange, int row, int col)
        {
            Excel.Range cell = (Excel.Range)colRange[row, col];
            object value = cell.Value;

            if (value == null)
            {
                return null;
            }

            string cellValue = value.ToString();

            Marshal.ReleaseComObject(cell);

            if (cellValue != null)
            {
                cellValue = cellValue.Trim();
            }
            return cellValue;
        }

        public static string GetCellColor(Excel.Range colRange, int row, int col)
        {
            Excel.Range cell = (Excel.Range)colRange[row, col];
            string color = cell.Interior.Color.ToString();

            Marshal.ReleaseComObject(cell);
            return color;
        }

        public static int GetColumnNumber(string columnLetters)
        {
            char[] columnLetterArray = columnLetters.ToCharArray();

            int index = 0;

            for (int i = 0; i < columnLetterArray.Length; i++)
            {
                int exponent = columnLetterArray.Length - 1 - i;
                index += (columnLetterArray[i] - 'A' + 1) * (int)Math.Pow(26, exponent);
            }

            return index;
        }

        public static Tuple<Excel.Range, Excel.Range> GetTable(ComReleaser releaser, string absolutePath, int worksheetNumber = 1)
        {
            if (!File.Exists(absolutePath))
            {
                return null;
            }

            Excel.Application excelApplication = new Excel.Application();
            Excel.Application xlApp = releaser.Add(() => excelApplication);
            Excel.Workbooks wBooks = releaser.Add(() => xlApp.Workbooks);
            Excel.Workbook excelWorkbook = wBooks.Open(absolutePath);
            Excel.Workbook xlWorkbook = releaser.Add(() => excelWorkbook);
            Excel.Sheets workSheets = releaser.Add(() => xlApp.Worksheets);
            Excel.Worksheet worksheet = (Excel.Worksheet)workSheets[worksheetNumber];
            Excel.Worksheet cuWorksheet = releaser.Add(() => worksheet);
            Excel.Range rowRange = releaser.Add(() => cuWorksheet.Rows);
            Excel.Range columnRange = releaser.Add(() => cuWorksheet.Cells);
            return new Tuple<Excel.Range, Excel.Range>(rowRange, columnRange);
        }
    }
}
