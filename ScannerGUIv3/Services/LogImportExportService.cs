using System.Text;
using System.Text.Json;
using ScannerGUIv3.Definitions;
using Serilog;
//using Application = Microsoft.UI.Xaml.Application;

namespace ScannerGUIv3.Services;

public class LogImportExportService
{
    private const string DownloadRosterUrl = "https://prod-29.australiasoutheast.logic.azure.com:443/workflows/293376f5258e440588acf2deed6bbe93/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=dIh-XwF7tJQL-zKPEiqu8wY3HhuLg26zrLDHAjtFlC8";
    public const string EmailUrl = "https://prod-02.australiasoutheast.logic.azure.com:443/workflows/94e6d29eed054a53b89b8448102a3ead/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ufoRVxP9dOGh8OP5VtuwYZAW3n25kYV8HN9_L8qnlGw";
    public const string TeamsUrl = "https://prod-03.australiaeast.logic.azure.com:443/workflows/dcd41880b0ab49b5a56f15e03fa3fbd7/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=PtCPW3Ch2ijFDjByU6dx1pUx_u1splgZFLQ2qwk1pjs";

    public static async Task<bool> DownloadExcelFileAsync(string filePath, string fileName)
    {
        Log.Information("Attempt Download Excel");
        // Validate inputs
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException(@"File path cannot be null or empty.", nameof(filePath));
        }

        if (string.IsNullOrWhiteSpace(fileName))
        {
            throw new ArgumentException(@"File name cannot be null or empty.", nameof(fileName));
        }

        //Console.WriteLine(filePath);
        // Ensure the directory exists
        Directory.CreateDirectory(filePath);

        try
        {
            using var client = new HttpClient();
            // Send GET request and get the response stream
            using var response = await client.GetAsync(DownloadRosterUrl, HttpCompletionOption.ResponseContentRead);
            response.EnsureSuccessStatusCode(); // Throws if not 200 OK

            // Combine the path and filename
            var fullFilePath = Path.Combine(filePath, fileName + ".xlsx");

            // Save the file content to disk
            await using (var contentStream = await response.Content.ReadAsStreamAsync())
            await using (var fileStream = new FileStream(fullFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                await contentStream.CopyToAsync(fileStream);
            }
            Log.Verbose($"Downloaded: {fullFilePath}");
            return true;
        }
        catch (HttpRequestException ex)
        {
            throw new HttpRequestException($"Failed to download the file: {ex.Message}", ex);
        }
    }

    public static string ExportShiftLog(Dictionary<int, Employee> employeeDict, string shiftType)
    {
        var csvBuilder = new StringBuilder();
        foreach (var employee in employeeDict.Values)
        {
            if (employee.ShiftType != shiftType)
            {
                continue;
            }

            var line = $"{employee.Name}, {employee.EmployeeNumber}, {employee.FormattedSignInTime}, {employee.FormattedSignOutTime}<br>";
            csvBuilder.AppendLine(line);
        }
        return csvBuilder.ToString();
    }

    public static async Task<bool> ExportToWeb(string url, object payload)
    {
        try
        {
            using var client = new HttpClient();
            var json = JsonSerializer.Serialize(payload);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var response = await client.PostAsync(url, content);
            return response.IsSuccessStatusCode;
        }
        catch (Exception ex)
        {
            // Log the error if needed; for now, return false to indicate failure
            return false;
        }
    }

    public static async Task SaveEmployeeDictionaryAsync(string filePath)
    {
        try
        {
            var json = JsonSerializer.Serialize(App.EmployeeDict);
            await File.WriteAllTextAsync(filePath, json);
            Log.Information("Employee Dictionary Saved: {FilePath}", filePath);
        }
        catch (Exception ex)
        {
            //Log.Error(ex, "Employee Dictionary NOT Saved {FilePath}", filePath);
            Log.Error(ex, "Employee Dictionary NOT Saved");
        }
    }

    public static async Task<Dictionary<int, Employee>> LoadEmployeeDictionaryAsync(string filePath)
    {
        try
        {
            var json = await File.ReadAllTextAsync(filePath);
            var employeeDict = JsonSerializer.Deserialize<Dictionary<int, Employee>>(json);
            Log.Information("Employee Dictionary Loaded: {FilePath}", filePath);
            return employeeDict ?? new Dictionary<int, Employee>();
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Employee Dictionary NOT Loaded {FilePath}", filePath);
            return new Dictionary<int, Employee>();
        }
    }

    public static string ExportShiftLogWithSignInStatus(Dictionary<int, Employee> employeeDict, string shiftType)
    {
        const string singleline =       "+---------------------------------------------------------------------------+";
        const string generaltitle =     "| Name                         | ID       | Sign In        | Sign Out       |";
        const string signedintitle =    "|                                 Signed In                                 |";
        const string notsignedintitle = "|                               Not Signed In                               |";
        const string tabledivider =     "+------------------------------+----------+----------------+----------------+";

        var csvBuilder = new StringBuilder();
        csvBuilder.AppendLine(singleline);
        csvBuilder.AppendLine(signedintitle);
        csvBuilder.AppendLine(tabledivider);
        csvBuilder.AppendLine(generaltitle);
        csvBuilder.AppendLine(tabledivider);

        // Employees with SignIn times
        foreach (var employee in employeeDict.Values)
        {
            if (employee.ShiftType == shiftType && employee.SignInTime.HasValue)
            {
                csvBuilder.AppendLine(employee.ToAsciiTableRow());
            }
        }

        csvBuilder.AppendLine(tabledivider);
        csvBuilder.AppendLine(notsignedintitle);
        csvBuilder.AppendLine(tabledivider);

        // Employees without SignIn times
        foreach (var employee in employeeDict.Values)
        {
            if (employee.ShiftType == shiftType && !employee.SignInTime.HasValue)
            {
                csvBuilder.AppendLine(employee.ToAsciiTableRow());
            }
        }

        csvBuilder.AppendLine(tabledivider);
        return csvBuilder.ToString();
    }

    // Retired
    //public static string ExportDayShiftLog(Dictionary<int, Employee> employeeDict)
    //{
    //    var csvBuilder = new StringBuilder();
    //    foreach (var employee in employeeDict.Values)
    //    {
    //        if (employee.ShiftType != "DS")
    //        {
    //            continue;
    //        }

    //        var line = $"{employee.Name}, {employee.EmployeeNumber}, {employee.FormattedSignInTime}, {employee.FormattedSignOutTime}<br>";
    //        csvBuilder.AppendLine(line);
    //    }
    //    return csvBuilder.ToString();
    //}

    //public static string ExportNightShiftLog(Dictionary<int, Employee> employeeDict)
    //{
    //    var csvBuilder = new StringBuilder();
    //    foreach (var employee in employeeDict.Values)
    //    {
    //        if (employee.ShiftType == "NS")
    //        {
    //            var line = $"{employee.Name}, {employee.EmployeeNumber}, {employee.FormattedSignInTime}, {employee.FormattedSignOutTime}<br>";
    //            csvBuilder.AppendLine(line);
    //        }
    //    }
    //    return csvBuilder.ToString();
    //}
}