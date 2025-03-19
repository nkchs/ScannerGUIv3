using System.Text;
using System.Text.Json;
using ScannerGUIv3.Definitions;
using Serilog;

namespace ScannerGUIv3.Services;

public class LogImportExportService
{
    private const string DownloadRosterUrl = "https://prod-29.australiasoutheast.logic.azure.com:443/workflows/293376f5258e440588acf2deed6bbe93/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=dIh-XwF7tJQL-zKPEiqu8wY3HhuLg26zrLDHAjtFlC8";
    public const string EmailUrl = "https://prod-02.australiasoutheast.logic.azure.com:443/workflows/94e6d29eed054a53b89b8448102a3ead/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ufoRVxP9dOGh8OP5VtuwYZAW3n25kYV8HN9_L8qnlGw";
    public const string TeamsUrl = "https://prod-03.australiaeast.logic.azure.com:443/workflows/dcd41880b0ab49b5a56f15e03fa3fbd7/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=PtCPW3Ch2ijFDjByU6dx1pUx_u1splgZFLQ2qwk1pjs";
    private const string TeamsTableUrl = "https://prod-31.australiaeast.logic.azure.com:443/workflows/212c98481a9642aba8db911f9a4b230a/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Ct06peZb9klgAGg0pMNihNl9vhydcyVLbhZagdrUMLk";

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
        const string singleline =       "+---------------------------------------------------------------------+";
        const string generaltitle =     "| Name                         | ID       | Sign In     | Sign Out    |";
        const string signedintitle =    "|                              Signed In                              |";
        const string notsignedintitle = "|                            Not Signed In                            |";
        const string tabledivider =     "+------------------------------+----------+-------------+-------------+";

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


    private static object[] GetEmployeeRows(Dictionary<int, Employee> employeeDict, string shiftType)
    {
        var rows = new List<object>
                {
                    new
                    {
                        type = "TableRow",
                        cells = new object[]
                        {
                            new
                            {
                                type = "TableCell",
                                items = new object[] { new { type = "TextBlock", weight = "Bolder", text = "Name" } }
                            },
                            new
                            {
                                type = "TableCell",
                                items = new object[] { new { type = "TextBlock", weight = "Bolder", text = "ID" } }
                            },
                            new
                            {
                                type = "TableCell",
                                items = new object[] { new { type = "TextBlock", weight = "Bolder", text = "Sign In" } }
                            },
                            new
                            {
                                type = "TableCell",
                                items = new object[] { new { type = "TextBlock", weight = "Bolder", text = "Sign Out" } }
                            }
                        }
                    }
                };

        foreach (var employee in employeeDict.Values)
        {
            if (employee.SignInTime.HasValue && employee.ShiftType == shiftType)
            {
                rows.Add(new
                {
                    type = "TableRow",
                    cells = new object[]
                    {
                        new
                        {
                            type = "TableCell",
                            items = new object[] { new { type = "TextBlock", text = employee.Name } }
                        },
                        new
                        {
                            type = "TableCell",
                            items = new object[] { new { type = "TextBlock", text = employee.EmployeeNumber.ToString() } }
                        },
                        new
                        {
                            type = "TableCell",
                            items = new object[] { new { type = "TextBlock", text = employee.FormattedSignInTime } }
                        },
                        new
                        {
                            type = "TableCell",
                            items = new object[] { new { type = "TextBlock", text = employee.FormattedSignOutTime ?? "No Sign Out" } }
                        }
                    }
                });
            }
        }

        return rows.ToArray();
    }


    public static async Task ExportAdaptiveCardFromTemplateAsync(string url, Dictionary<int, Employee> employeeDict, string shiftType)
    {
        // Define the JSON structure using anonymous objects with explicit array typing
        var teamsMessage = new
        {
            type = "message",
            attachments = new object[]
            {
                new
                {
                    contentType = "application/vnd.microsoft.card.adaptive",
                    content = new
                    {
                        type = "AdaptiveCard",
                        schema = "http://adaptivecards.io/schemas/adaptive-card.json",
                        version = "1.6",
                        msteams = new { width = "Full" },
                        body = new object[]
                        {
                            new
                            {
                                type = "TextBlock",
                                size = "Medium",
                                weight = "Bolder",
                                text = "Attendance Report"
                            },
                            new
                            {
                                type = "TextBlock",
                                text = "Signed In",
                                wrap = true
                            },
                            new
                            {
                                type = "Table",
                                columns = new object[]
                                {
                                    new { width = 2 },
                                    new { width = 1 },
                                    new { width = 1 },
                                    new { width = 1 }
                                },
                                rows = GetEmployeeRows(employeeDict, shiftType)
                            }
                        }
                    }
                }
            }
        };

        // Serialize to JSON using System.Text.Json
        var json = JsonSerializer.Serialize(teamsMessage, new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = null  // Ensures property names match exactly as defined
        });

        // Send to the URL
        try
        {
            using var client = new HttpClient();
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var response = await client.PostAsync(TeamsTableUrl, content);
            response.EnsureSuccessStatusCode();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error posting to Teams: {ex.Message}");
            throw;
        }
    }
}


//public static string ExportAdaptiveCardFromTemplate(string url, Dictionary<int, Employee> employeeDict)
//{
//    // Define the JSON structure using anonymous objects with explicit array typing
//    var teamsMessage = new
//    {
//        type = "message",
//        attachments = new object[]
//        {
//                    new
//                    {
//                        contentType = "application/vnd.microsoft.card.adaptive",
//                        content = new
//                        {
//                            type = "AdaptiveCard",
//                            schema = "http://adaptivecards.io/schemas/adaptive-card.json",
//                            version = "1.6",
//                            msteams = new { width = "Full" },
//                            body = new object[]
//                            {
//                                new
//                                {
//                                    type = "TextBlock",
//                                    size = "Medium",
//                                    weight = "Bolder",
//                                    text = "Attendance Report"
//                                },
//                                new
//                                {
//                                    type = "TextBlock",
//                                    text = "Signed In",
//                                    wrap = true
//                                },
//                                new
//                                {
//                                    type = "Table",
//                                    columns = new object[]
//                                    {
//                                        new { width = 2 },
//                                        new { width = 1 },
//                                        new { width = 1 },
//                                        new { width = 1 }
//                                    },
//                                    rows = GetEmployeeRows(employeeDict)
//                                }
//                            }
//                        }
//                    }
//        }
//    };

//    // Serialize to JSON using System.Text.Json
//    var json = JsonSerializer.Serialize(teamsMessage, new JsonSerializerOptions
//    {
//        WriteIndented = true,
//        PropertyNamingPolicy = null  // Ensures property names match exactly as defined
//    });

//    //Console.WriteLine(json);

//    // Send to the URL
//    try
//    {
//        using var client = new HttpClient();
//        var content = new StringContent(json, Encoding.UTF8, "application/json");
//        var response = client.PostAsync(TeamsTableUrl, content).Result;
//        Console.WriteLine(response);
//        response.EnsureSuccessStatusCode();
//    }
//    catch (Exception ex)
//    {
//        Console.WriteLine($"Error posting to Teams: {ex.Message}");
//        throw;
//    }

//    return json;
//}


//public static string ExportShiftLogToJson(Dictionary<int, Employee> employeeDict, string shiftType)
//{
//    var jsonBuilder = new StringBuilder();

//    jsonBuilder.Append(
//        "{\r\n  \"type\": \"message\",\r\n  \"attachments\": [\r\n    {\r\n      \"contentType\": \"application/vnd.microsoft.card.adaptive\",\r\n      \"content\": {\r\n        \"msteams\": {\r\n          \"width\": \"Full\"\r\n        },\r\n        \"$schema\": \"http://adaptivecards.io/schemas/adaptive-card.json\",\r\n        \"type\": \"AdaptiveCard\",\r\n        \"version\": \"1.5\",\r\n        \"body\": [\r\n          {\r\n            \"type\": \"TextBlock\",\r\n            \"size\": \"Medium\",\r\n            \"weight\": \"Bolder\",\r\n            \"text\": \"Attendance Report\"\r\n          },\r\n          {\r\n            \"type\": \"TextBlock\",\r\n            \"text\": \"Signed In\",\r\n            \"wrap\": true\r\n          },\r\n          {\r\n            \"type\": \"Table\",\r\n            \"columns\": [\r\n              { \"width\": 2 },  // Increased width for Name column\r\n              { \"width\": 1 },\r\n              { \"width\": 1 },\r\n              { \"width\": 1 }\r\n            ],\r\n            \"rows\": ");


//    jsonBuilder.AppendLine("[");

//    // Employees with SignIn times
//    foreach (var employee in employeeDict.Values)
//    {
//        if (employee.ShiftType == shiftType && employee.SignInTime.HasValue)
//        {
//            jsonBuilder.AppendLine(employee.ToJsonTableRow() + ",");
//        }
//    }

//    // Remove the last comma and close the JSON array
//    if (jsonBuilder.Length > 1)
//    {
//        jsonBuilder.Length--; // Remove the last comma
//    }
//    jsonBuilder.AppendLine("]");

//    return jsonBuilder.ToString();
//}

//public static string CreateAdaptiveCard(Dictionary<int, Employee> employeeDict, string shiftType)
//{
//    var signedInEmployees = employeeDict.Values
//        .Where(e => e.ShiftType == shiftType && e.SignInTime.HasValue)
//        .Select(e => new
//        {
//            type = "TableRow",
//            cells = new[]
//            {
//                new { type = "TableCell", items = new[] { new { type = "TextBlock", text = e.Name } } },
//                new { type = "TableCell", items = new[] { new { type = "TextBlock", text = e.EmployeeNumber.ToString() } } },
//                new { type = "TableCell", items = new[] { new { type = "TextBlock", text = e.SignInTime?.ToString("HH:mm") ?? "No Sign In" } } },
//                new { type = "TableCell", items = new[] { new { type = "TextBlock", text = e.SignOutTime?.ToString("HH:mm") ?? "No Sign Out" } } }
//            }
//        })
//        .ToArray();

//    var notSignedInEmployees = employeeDict.Values
//        .Where(e => e.ShiftType == shiftType && !e.SignInTime.HasValue)
//        .Select(e => new
//        {
//            type = "TableRow",
//            cells = new[]
//            {
//                new { type = "TableCell", items = new[] { new { type = "TextBlock", text = e.Name } } },
//                new { type = "TableCell", items = new[] { new { type = "TextBlock", text = e.EmployeeNumber.ToString() } } },
//                new { type = "TableCell", items = new[] { new { type = "TextBlock", text = "No Sign In" } } },
//                new { type = "TableCell", items = new[] { new { type = "TextBlock", text = "No Sign Out" } } }
//            }
//        })
//        .ToArray();

//    var card = new
//    {
//        type = "message",
//        attachments = new[]
//        {
//            new
//            {
//                contentType = "application/vnd.microsoft.card.adaptive",
//                content = new
//                {
//                    msteams = new { width = "Full" },
//                    schema = "http://adaptivecards.io/schemas/adaptive-card.json",
//                    type = "AdaptiveCard",
//                    version = "1.5",
//                    body = new object[]
//                    {
//                        new { type = "TextBlock", size = "Medium", weight = "Bolder", text = "Attendance Report" },
//                        new { type = "TextBlock", text = "Signed In", wrap = true },
//                        new
//                        {
//                            type = "Table",
//                            columns = new[]
//                            {
//                                new { width = 2 },
//                                new { width = 1 },
//                                new { width = 1 },
//                                new { width = 1 }
//                            },
//                            rows = signedInEmployees
//                        },
//                        new { type = "TextBlock", text = "Not Signed In", wrap = true },
//                        new
//                        {
//                            type = "Table",
//                            columns = new[]
//                            {
//                                new { width = 2 },
//                                new { width = 1 },
//                                new { width = 1 },
//                                new { width = 1 }
//                            },
//                            rows = notSignedInEmployees
//                        }
//                    }
//                }
//            }
//        }
//    };

//    return JsonSerializer.Serialize(card);
//}

//public static async Task<bool> SendAdaptiveCardAsync(Dictionary<int, Employee> employeeDict, string shiftType, string url)
//{
//    try
//    {
//        var adaptiveCardJson = CreateAdaptiveCard(employeeDict, shiftType);
//        using var client = new HttpClient();
//        var content = new StringContent(adaptiveCardJson, Encoding.UTF8, "application/json");
//        var response = await client.PostAsync(url, content);
//        return response.IsSuccessStatusCode;
//    }
//    catch (Exception ex)
//    {
//        Log.Error(ex, "Failed to send adaptive card");
//        return false;
//    }
//}



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
//}