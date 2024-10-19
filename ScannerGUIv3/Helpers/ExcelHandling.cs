using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ScannerGUIv3.Helpers;
internal class ExcelHandling
{
    public static void ExcelLoader()
    {
        var excel = new Application
        {
            Visible = false,
            // Visible = true,
        };
        Console.WriteLine("ExcelHandling.cs");
        Workbook excelWorkbook = excel.Workbooks.Open(@"C:\Users\ChaseN\ResourceOnSite_20240620043001.xlsx");

        Worksheet excelWorksheet = (Worksheet)excelWorkbook.Sheets[3];
        Range excelRange = excelWorksheet.UsedRange;

        int maxRow = excelRange.Rows.Count;
        int maxCol = excelRange.Columns.Count;
        Console.WriteLine("Rows: " + maxRow);
        Console.WriteLine("Columns: " + maxCol);
    }
}
