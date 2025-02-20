using System.Text;
using ScannerGUIv3;
using ScannerGUIv3.Models;
using ScannerGUIv3.Definitions;

namespace ScannerGUIv3.Services
{
    public class LogImportExportService
    {
        public static string ExportDayShiftLog(Dictionary<int, Employee> employeeDict)
        {
            var csvBuilder = new StringBuilder();
            //csvBuilder.AppendLine("EmployeeNumber, Name, FormattedSignInTime, FormattedSignOutTime");

            foreach (var employee in employeeDict.Values)
            {
                //Console.WriteLine(employee.Name + " " + employee.ShiftType);
                if (employee.ShiftType == "DS")
                {
                    var line = $"{employee.EmployeeNumber},{employee.Name},{employee.FormattedSignInTime},{employee.FormattedSignOutTime}";
                    csvBuilder.AppendLine(line);
                }
            }

            return csvBuilder.ToString();
        }



    }
}
