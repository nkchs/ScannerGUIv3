//using Microsoft.Office.Interop.Excel;
using ScannerGUIv3.Definitions;
using Windows.Media.Streaming.Adaptive;
//using Range = Microsoft.Office.Interop.Excel.Range;

namespace ScannerGUIv3.Helpers;

//public static class ExcelHandling
//{
//    //public static void ExcelGetter()
//    //{
//    //    var excel = new Application
//    //    {
//    //        Visible = false,
//    //        // Visible = true,
//    //    };
//    //    Console.WriteLine("ExcelHandling.cs");
//    //    Workbook excelWorkbook = excel.Workbooks.Open(@"C:\Users\ChaseN\ResourceOnSite_20240620043001.xlsx");

//    //    Worksheet excelWorksheet = (Worksheet)excelWorkbook.Sheets[3];
//    //    Range excelRange = excelWorksheet.UsedRange;

//    //    int maxRow = excelRange.Rows.Count;
//    //    int maxCol = excelRange.Columns.Count;
//    //    Console.WriteLine("Rows: " + maxRow);
//    //    Console.WriteLine("Columns: " + maxCol);
//    //}

//    //public static string ExcelPopulator(Dictionary<int, Employee> _employeedict)
//    //{
//    //    // 
//    //    // Cycle through the employee dict.
//    //    // If the shift is DS then populate the DS page... same for NS.
//    //    // To populate the CORRECT page you'll need DateTime.now(), then Week, Day of Week.
        
//    //    // If SignInTime exists, populate.
//    //    // If SignOutTime exists, populate.
//    //    // If SignOutTime exists, but SignInTime doesn't, populate SignOutTime, & set SignInTime to DNSI (Did not Sign In).

//    //    // If SignInTime exists & SignOutTime doesn't exist && it's after 6:30PM or AM that is after the shift start, then populate SignInTime and set SignOutTime to DNSO (Did Not Sign Out)
//    //    // If SignInTime == null && SignOutTime == null && Datetime.now is beyond the shift end.

//    //    //EmployeeDict.
//    //    return "Test";
//    //}
//}