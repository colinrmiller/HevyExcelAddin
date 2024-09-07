using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace HevyAddIn
{
    public class Helpers
    {
        /*
         * Format Columns to avoid Word wrap
         * 
         * TODO: column width needs to be set manually and calculated from the length of the workout
         *        - this should be performed after assigning fixed columns to each excersize
         */
        public static void FormatColumns()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet worksheet = excelApp.ActiveSheet as Excel.Worksheet;

            if (worksheet == null)
            {
                return;
            }

            int lastColumn = FindLastColumnWithData(worksheet);

            // Loop through each column to auto-fit and then adjust
            for (int i = 1; i <= lastColumn; i++)
            {
                Excel.Range column = worksheet.Columns[i];

                column.AutoFit();

                //column.WrapText = false;
            }
        }
        public static double ConvertKgToLbs(double? kg)
        {
            const double kgToLbsFactor = 2.20462;

            if (kg == null) { return 0; }
            double lbs = (double)(kg * kgToLbsFactor);

            return Math.Round(lbs);
        }

        public static string FormatExcerciseTitle(string excerciseTitle)
        {
            excerciseTitle = excerciseTitle.Replace(" - ", "\n");
            excerciseTitle = excerciseTitle.Replace("-", "\n");
            excerciseTitle = excerciseTitle.Replace(" (", "\n(");
            excerciseTitle = excerciseTitle.Replace("(", "\n(");
            excerciseTitle = excerciseTitle.Replace("\n\n", "\n");
            return excerciseTitle;
        }

        public static int FindLastColumnWithData(Excel.Worksheet worksheet)
        {
            int lastColIndex = 1;
            int lastRow = FindLastRowWithData(worksheet);
            // Iterate over each row
            for (int i = 1; i <= lastRow; i++)
            {
                // Find the last used cell in this row
                Excel.Range lastCellInRow = worksheet.Cells[i, worksheet.Columns.Count].End(Excel.XlDirection.xlToLeft);

                // Update lastColIndex if this cell's column index is greater
                if (lastCellInRow.Column > lastColIndex)
                {
                    lastColIndex = lastCellInRow.Column;
                }
            }

            return lastColIndex;

        }

        public static int FindLastRowWithData(Excel.Worksheet worksheet)
        {
            return worksheet.Cells[1000000, 1].End(Excel.XlDirection.xlUp).Row;
        }

        public static string[] RemoveDuplicates(string[] inputArray)
        {
            HashSet<string> seen = new HashSet<string>();
            List<string> result = new List<string>();

            foreach (string item in inputArray)
            {
                if (seen.Add(item)) 
                {
                    result.Add(item); 
                }
            }

            return result.ToArray();
        }

        public static List<string> RemoveDuplicates(List<string> inputArray)
        {
            HashSet<string> seen = new HashSet<string>();
            List<string> result = new List<string>();

            foreach (string item in inputArray)
            {
                if (seen.Add(item))
                {
                    result.Add(item);
                }
            }

            return result;
        }
    }
}
