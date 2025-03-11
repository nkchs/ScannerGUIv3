using System.Text;
using System.Text.Json;
using ScannerGUIv3.Definitions;
using Serilog;
//using Application = Microsoft.UI.Xaml.Application;

namespace ScannerGUIv3.Services;

public class LogImportExportService
{
    private static readonly string downloadRosterUrl = "https://prod-29.australiasoutheast.logic.azure.com:443/workflows/293376f5258e440588acf2deed6bbe93/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=dIh-XwF7tJQL-zKPEiqu8wY3HhuLg26zrLDHAjtFlC8";
    public static readonly string emailUrl = "https://prod-02.australiasoutheast.logic.azure.com:443/workflows/94e6d29eed054a53b89b8448102a3ead/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ufoRVxP9dOGh8OP5VtuwYZAW3n25kYV8HN9_L8qnlGw";
    public static readonly string teamsUrl = "https://prod-08.australiasoutheast.logic.azure.com:443/workflows/ffd31ea3fab043d088f02cfbc959548e/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=KkGwF2ZPk7lbUde9u2SHXVoDnVbxuLJcdHAQ5KNHjKg";

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
            using var response = await client.GetAsync(downloadRosterUrl, HttpCompletionOption.ResponseContentRead);
            response.EnsureSuccessStatusCode(); // Throws if not 200 OK

            // Combine the path and filename
            var fullFilePath = Path.Combine(filePath, fileName + ".xlsx");

            // Save the file content to disk
            await using (Stream contentStream = await response.Content.ReadAsStreamAsync())
            await using (FileStream fileStream = new FileStream(fullFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
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

    public static string ExportDayShiftLog(Dictionary<int, Employee> employeeDict)
    {
        var csvBuilder = new StringBuilder();
        foreach (var employee in employeeDict.Values)
        {
            if (employee.ShiftType == "DS")
            {
                var line = $"{employee.Name}, {employee.EmployeeNumber}, {employee.FormattedSignInTime}, {employee.FormattedSignOutTime}<br>";
                csvBuilder.AppendLine(line);
            }
        }
        return csvBuilder.ToString();
    }

    public static string ExportNightShiftLog(Dictionary<int, Employee> employeeDict)
    {
        var csvBuilder = new StringBuilder();
        foreach (var employee in employeeDict.Values)
        {
            if (employee.ShiftType == "NS")
            {
                var line = $"{employee.Name}, {employee.EmployeeNumber}, {employee.FormattedSignInTime}, {employee.FormattedSignOutTime}<br>";
                csvBuilder.AppendLine(line);
            }
        }
        return csvBuilder.ToString();
    }

    public static string ExportShiftLog(Dictionary<int, Employee> employeeDict, String ShiftType)
    {
        var csvBuilder = new StringBuilder();
        foreach (var employee in employeeDict.Values)
        {
            if (employee.ShiftType == ShiftType)
            {
                var line = $"{employee.Name}, {employee.EmployeeNumber}, {employee.FormattedSignInTime}, {employee.FormattedSignOutTime}<br>";
                csvBuilder.AppendLine(line);
            }
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


}