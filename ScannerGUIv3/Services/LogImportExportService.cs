using System.Text;
using System.Text.Json;
//using Newtonsoft.Json;
using ScannerGUIv3.Definitions;
using Serilog;

namespace ScannerGUIv3.Services;

public class LogImportExportService
{
    //private const string DownloadRosterUrl = "https://prod-29.australiasoutheast.logic.azure.com:443/workflows/293376f5258e440588acf2deed6bbe93/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=dIh-XwF7tJQL-zKPEiqu8wY3HhuLg26zrLDHAjtFlC8";
    private const string DownloadRosterUrl = "https://prod-31.australiaeast.logic.azure.com:443/workflows/632911e333f54280b5f23c1fdad9039b/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=PB6P68kCxrWk2PYPYqenD1hY2fHPiWQhUPKfJ0x6vbc";
    public const string EmailUrl = "https://prod-02.australiasoutheast.logic.azure.com:443/workflows/94e6d29eed054a53b89b8448102a3ead/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ufoRVxP9dOGh8OP5VtuwYZAW3n25kYV8HN9_L8qnlGw";
    public const string TeamsUrl = "https://prod-03.australiaeast.logic.azure.com:443/workflows/dcd41880b0ab49b5a56f15e03fa3fbd7/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=PtCPW3Ch2ijFDjByU6dx1pUx_u1splgZFLQ2qwk1pjs";
    private const string TeamsTableUrl = "https://prod-31.australiaeast.logic.azure.com:443/workflows/212c98481a9642aba8db911f9a4b230a/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Ct06peZb9klgAGg0pMNihNl9vhydcyVLbhZagdrUMLk";
    private const string RosterDateUrl =
        "https://prod-39.australiasoutheast.logic.azure.com:443/workflows/77439615022643799f39a62f6d6704b6/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=tBbbMYRU4lAzKJZsVILQnA1BT4WooIvvWekVziKEhNw";
    private const string EmailTableUrl = "https://prod-19.australiaeast.logic.azure.com:443/workflows/512e71742dcc42a18aadc445eaad070d/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=s5AoSnzv_rwQgYo-b5uucsmHcTdB-QlOoyIyRnu20LU";

    public static async Task<bool> GetRosterDate()
    {
        try
        {
            using var client = new HttpClient();
            using var response = await client.GetAsync(RosterDateUrl, HttpCompletionOption.ResponseContentRead);

            // Ensure the response is successful
            response.EnsureSuccessStatusCode();
            Log.Debug("Response received. Status Code: {StatusCode}", response.StatusCode);

            // Read the response as a string
            var responseBody = await response.Content.ReadAsStringAsync();

            // Trim any extra whitespace or newline characters from the response
            responseBody = responseBody.Trim();

            // Parse the response into a DateTime object
            if (DateTime.TryParse(responseBody, out var responseDateTime))
            {
                // Since response is already in WA time, get today's date in WA time
                TimeZoneInfo waTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Australia/Perth");
                var waToday = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, waTimeZone).Date;

                // Compare dates (ignoring time)
                if (responseDateTime.Date == waToday)
                {
                    Log.Debug($"Roster Has Been Updated @ {responseDateTime}");
                    return true;
                }
            }
            Log.Error("Roster Has NOT Been Updated");
            return false;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
            return false;
        }
    }

    public static async Task<bool> DownloadExcelFileAsync(string filePath, string fileName)
    {
        Log.Information("Attempting to download Excel file to {FilePath} with file name {FileName}", filePath, fileName);
        try
        {
            // Validate inputs
            if (string.IsNullOrWhiteSpace(filePath))
            {
                Log.Error("Validation failed: File path is null or empty.");
                throw new ArgumentException(@"File path cannot be null or empty.", nameof(filePath));
            }

            if (string.IsNullOrWhiteSpace(fileName))
            {
                Log.Error("Validation failed: File name is null or empty.");
                throw new ArgumentException(@"File name cannot be null or empty.", nameof(fileName));
            }

            // Ensure the directory exists
            Log.Debug("Ensuring directory exists at path: {FilePath}", filePath);
            Directory.CreateDirectory(filePath);

            using var client = new HttpClient();

            // Log request initiation
            Log.Debug("Sending GET request to {DownloadRosterUrl} with HttpCompletionOption.ResponseContentRead", DownloadRosterUrl);

            using var response = await client.GetAsync(DownloadRosterUrl, HttpCompletionOption.ResponseContentRead);

            // Log response details
            Log.Debug("Response received. Status Code: {StatusCode}", response.StatusCode);

            response.EnsureSuccessStatusCode(); // Throws if not 200 OK

            // Combine the path and filename
            var fullFilePath = Path.Combine(filePath, fileName + ".xlsx");
            Log.Debug("Full file path resolved to: {FullFilePath}", fullFilePath);

            // Save the file content to disk
            Log.Debug("Starting to copy content stream to disk.");
            await using (var contentStream = await response.Content.ReadAsStreamAsync())
            await using (var fileStream = new FileStream(fullFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                await contentStream.CopyToAsync(fileStream);
            }

            Log.Information("Excel file successfully downloaded to {FullFilePath}", fullFilePath);
            return true;
        }
        catch (HttpRequestException ex)
        {
            Log.Error(ex, "HTTP error occurred while downloading file: {Message}", ex.Message);
            return false;
        }
        catch (IOException ex)
        {
            Log.Error(ex, "File I/O error occurred while saving the file: {Message}", ex.Message);
            return false;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Unexpected error occurred: {Message}", ex.Message);
            return false;
        }
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
            return false;
        }
    }

    public static async Task SaveEmployeeDictionaryAsync(string filePath)
    {
        try
        {
            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(App.EmployeeDict, options);
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

    private static object[] GetEmployeeRows(Dictionary<int, Employee> employeeDict, string shiftType, bool signedIn = true)
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
            if (employee.SignInTime.HasValue == signedIn && employee.ShiftType == shiftType)
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
                            items = new object[]
                                { new { type = "TextBlock", text = employee.EmployeeNumber.ToString() } }
                        },
                        new
                        {
                            type = "TableCell",
                            items = new object[] { new { type = "TextBlock", text = employee.FormattedSignInTime } }
                        },
                        new
                        {
                            type = "TableCell",
                            items = new object[]
                            {
                                new { type = "TextBlock", text = employee.FormattedSignOutTime ?? "No Sign Out" }
                            }
                        }
                    }
                });
            }
        }

        return rows.ToArray();
    }

    public static async Task<bool> SendEmployeeDataAsync(string email, string shiftType)
    {
        // Prepare the JSON string
        var jsonData = PrepareJsonForEmail(email, shiftType);

        // Create HttpClient instance
        using var client = new HttpClient();

        // Create StringContent with JSON data
        var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

        try
        {
            // Send POST request to EmailTableUrl
            var response = await client.PostAsync(EmailTableUrl, content);

            // Ensure the request was successful and return true
            response.EnsureSuccessStatusCode();
            return true;
        }
        catch (HttpRequestException)
        {
            // Return false on HTTP-related errors instead of throwing
            return false;
        }
    }

    private static string PrepareJsonForEmail(string email, string shiftType)
    {
        var signedInList = GetEmployeeDataForJson(shiftType, true);  // Changed "DS" to shiftType
        var notSignedInList = GetEmployeeDataForJson(shiftType, false);

        var message = new
        {
            email = email,
            signedIn = signedInList,    // Already an array, no need for ToArray()
            notSignedIn = notSignedInList  // Already an array, no need for ToArray()
        };
        return JsonSerializer.Serialize(message);  // Assuming you want to return JSON string
    }

    private static object[] GetEmployeeDataForJson(string shiftType, bool signedIn = true)
    {
        var employeeList = new List<object>();

        foreach (var employee in App.EmployeeDict.Values.Where(employee => employee.SignInTime.HasValue == signedIn && employee.ShiftType == shiftType))
        {
            employeeList.Add(new
            {
                employeeID = employee.EmployeeNumber.ToString(),
                name = employee.Name,
                signIn = employee.FormattedSignInTime,
                signOut = employee.FormattedSignOutTime ?? "No Sign Out"
            });
        }

        return employeeList.ToArray();
    }

    public static string PrepareJsonMessageForPowerAutomate(string shiftType)
    {
        var signedInRows = GetEmployeeRows(App.EmployeeDict, shiftType);
        var notSignedInRows = GetEmployeeRows(App.EmployeeDict, shiftType, false);

        var message = new
        {
            signedIn = signedInRows,
            notSignedIn = notSignedInRows
        };

        return System.Text.Json.JsonSerializer.Serialize(message, new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });
        ;
    }

    public static async Task<bool> ExportAdaptiveCardFromTemplateAsync(Dictionary<int, Employee> employeeDict, string shiftType, string message = "")
    {
        try
        {
            // Determine the shift title based on shiftType
            var shiftTitle = shiftType == "DS" ? "Day Shift" : shiftType == "NS" ? "Night Shift" : "Unknown Shift";

            // Log the start of the method
            Log.Information("ExportAdaptiveCardFromTemplateAsync started with shiftType: {ShiftType}", shiftType);

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
                                    text = $"Attendance Report - {shiftTitle}"
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
                                },
                                new
                                {
                                    type = "TextBlock",
                                    text = "Not Signed In",
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
                                    rows = GetEmployeeRows(employeeDict, shiftType, false)
                                }
                            }
                        }
                    }
                }
            };

            var firstAttachment = ((dynamic)teamsMessage).attachments[0];
            //var firstItemInAttachment = firstAttachment.content.body[0];

            // Calculate the number of bytes
            string firstAttachmentJson = JsonSerializer.Serialize(firstAttachment, new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = null // Ensures property names match exactly as defined));
            });

            var byteCount = Encoding.UTF8.GetByteCount(firstAttachmentJson);
            Log.Information("The JSON string is {ByteCount} bytes long.", byteCount);
            //var kiloByteCount = byteCount / 1024.0; // Use 1024.0 to ensure floating-point division
            //Log.Information("The JSON string is {KiloByteCount:F2} kB long.", kiloByteCount);
            
            // Serialize to JSON using System.Text.Json
            var json = JsonSerializer.Serialize(teamsMessage, new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = null // Ensures property names match exactly as defined
            });

            // Send to the URL
            using var client = new HttpClient();
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var response = await client.PostAsync(TeamsTableUrl, content);

            response.EnsureSuccessStatusCode(); // Throws an exception if the status code is not successful

            // Log the success of the operation
            Log.Information("Successfully posted the adaptive card to Teams for shiftType: {ShiftType}", shiftType);
            return true; // Operation succeeded
        }
        catch (Exception ex)
        {
            // Log the exception
            Log.Error(ex, "Error posting adaptive card to Teams for shiftType: {ShiftType}", shiftType);
            return false; // Indicate failure
        }
    }

    // RETIRED
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
}