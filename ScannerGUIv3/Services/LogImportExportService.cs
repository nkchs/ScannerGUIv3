using System.Text;
using ScannerGUIv3;
using ScannerGUIv3.Models;
using ScannerGUIv3.Definitions;

namespace ScannerGUIv3.Services
{
    public class LogImportExportService
    {
        public static readonly string teamsUrl = "https://prod-08.australiasoutheast.logic.azure.com:443/workflows/ffd31ea3fab043d088f02cfbc959548e/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=KkGwF2ZPk7lbUde9u2SHXVoDnVbxuLJcdHAQ5KNHjKg";
        public static readonly string emailUrl = "";

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
            //_ = ExportToWeb(csvBuilder.ToString());
            //_ = ExportToWeb(csvBuilder.ToString(), teamsUrl);
            //Console.WriteLine(teamsUrl);
            return csvBuilder.ToString();
        }

        public static async Task ExportToWeb(string message, string url)
        {
            //var url = "https://prod-08.australiasoutheast.logic.azure.com:443/workflows/ffd31ea3fab043d088f02cfbc959548e/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=KkGwF2ZPk7lbUde9u2SHXVoDnVbxuLJcdHAQ5KNHjKg";

            using var client = new HttpClient();
            var content = new StringContent($"{{\"message\":\"{message}\"}}", Encoding.UTF8, "application/json");
            var response = await client.PostAsync(url, content);

            if (response.IsSuccessStatusCode)
            {
                Console.WriteLine("Request sent successfully.");
            }
            else
            {
                Console.WriteLine($"Failed to send request. Status code: {response.StatusCode}");
            }
        }

        //private static async Task ExportToWeb(string message)
        //{
        //    var url = "https://prod-08.australiasoutheast.logic.azure.com:443/workflows/ffd31ea3fab043d088f02cfbc959548e/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=KkGwF2ZPk7lbUde9u2SHXVoDnVbxuLJcdHAQ5KNHjKg";
        //    using var client = new HttpClient();
        //    var content = new StringContent($"{{\"message\":\"{message}\"}}", Encoding.UTF8, "application/json");
        //    var response = await client.PostAsync(url, content);
        //    if (response.IsSuccessStatusCode)
        //    {
        //        Console.WriteLine("Request sent successfully.");
        //    }
        //    else
        //    {
        //        Console.WriteLine($"Failed to send request. Status code: {response.StatusCode}");
        //    }
        //}
    }
}
