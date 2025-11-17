using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using ScannerGUIv3.Definitions;
using Serilog;

namespace ScannerGUIv3.Services;

public class LogImportExportService
{
    private static readonly object _dictLock = new object(); // Shared with ExcelService
    //private const string DownloadRosterUrl = "https://prod-31.australiaeast.logic.azure.com:443/workflows/632911e333f54280b5f23c1fdad9039b/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=PB6P68kCxrWk2PYPYqenD1hY2fHPiWQhUPKfJ0x6vbc";
    //public const string EmailUrl = "https://prod-02.australiasoutheast.logic.azure.com:443/workflows/94e6d29eed054a53b89b8448102a3ead/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ufoRVxP9dOGh8OP5VtuwYZAW3n25kYV8HN9_L8qnlGw";
    private const string TeamsTableUrl = "https://prod-06.australiaeast.logic.azure.com:443/workflows/55864becb65844baa48749bf029985df/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=9VFdB_qY9-fdrAcS-lbiBqYKNZjaRZqBe3oQxvOiZ40";
    //private const string RosterDateUrl = "https://prod-39.australiasoutheast.logic.azure.com:443/workflows/77439615022643799f39a62f6d6704b6/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=tBbbMYRU4lAzKJZsVILQnA1BT4WooIvvWekVziKEhNw";
    private const string EmailTableUrl = "https://prod-19.australiaeast.logic.azure.com:443/workflows/512e71742dcc42a18aadc445eaad070d/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=s5AoSnzv_rwQgYo-b5uucsmHcTdB-QlOoyIyRnu20LU";
    public const string WorkforceJobUrl = "https://reportingtel.vixresources.com/api/external/saved-reports/FPM%20Roster%20Dataset%20SRF175%20Roster%20to%20Excel%20Today_Plus_14days";
    public const string bearerToken = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJ0ZWwucHJvZCIsImlhdCI6MTc1Njk3NTM4NSwiZXhwIjoxNzg4NTExMzg1LCJhdWQiOiJodHRwczovL3JlcG9ydGluZ3RlbC52aXhyZXNvdXJjZXMuY29tIiwiaXNzIjoiaW54c29mdHdhcmUuY29tIn0.b3laHNksViAmp_tsIcHfYevm4J501mtj1u_tLDZbgg4";
    public static string WorkforceReportDownloadUri = "";

    //public static async Task<bool> GetRosterDate()
    //{
    //    try
    //    {
    //        using var client = new HttpClient();
    //        using var response = await client.GetAsync(RosterDateUrl, HttpCompletionOption.ResponseContentRead);
    //        response.EnsureSuccessStatusCode();
    //        Log.Debug("Response received from RosterDateUrl. Status Code: {StatusCode}", response.StatusCode);

    //        var responseBody = await response.Content.ReadAsStringAsync();
    //        responseBody = responseBody.Trim();

    //        if (DateTime.TryParse(responseBody, out var responseDateTime))
    //        {
    //            TimeZoneInfo waTimeZone = TimeZoneInfo.FindSystemTimeZoneById("W. Australia Standard Time");
    //            var waToday = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, waTimeZone).Date;
    //            if (responseDateTime.Date == waToday)
    //            {
    //                Log.Debug("Roster updated at {DateTime}", responseDateTime);
    //                return true;
    //            }
    //        }
    //        Log.Error("Roster not updated. Response: {ResponseBody}", responseBody);
    //        return false;
    //    }
    //    catch (Exception ex)
    //    {
    //        Log.Error(ex, "Failed to get roster date");
    //        return false;
    //    }
    //}

    public static async Task<bool> DownloadExcelFileAsync(string filePath, string fileName, string url)
    {
        Log.Information("Attempting to download Excel file from {Url} to {FilePath} with file name {FileName}", url, filePath, fileName);
        try
        {
            // Validate inputs
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));
            }

            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("File name cannot be null or empty.", nameof(fileName));
            if (string.IsNullOrWhiteSpace(url))
                throw new ArgumentException("URL cannot be null or empty.", nameof(url));

            Log.Debug("Ensuring directory exists at path: {FilePath}", filePath);
            Directory.CreateDirectory(filePath);

            var fullFilePath = Path.Combine(filePath, fileName + ".xlsx");
            var tempFilePath = Path.Combine(filePath, $"Roster_{Guid.NewGuid()}.xlsx");

            const int maxRetries = 3;
            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    using var client = new HttpClient();
                    Log.Debug("Sending GET request to {Url} with HttpCompletionOption.ResponseContentRead (Attempt {Attempt})", url, attempt);
                    using var response = await client.GetAsync(url, HttpCompletionOption.ResponseContentRead);
                    response.EnsureSuccessStatusCode();
                    Log.Debug("Response received. Status Code: {StatusCode}", response.StatusCode);

                    Log.Debug("Saving to temporary file: {TempFilePath}", tempFilePath);
                    await using (var contentStream = await response.Content.ReadAsStreamAsync())
                    await using (var fileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        await contentStream.CopyToAsync(fileStream);
                    }

                    if (File.Exists(fullFilePath))
                    {
                        Log.Debug("Deleting existing file: {FullFilePath}", fullFilePath);
                        File.Delete(fullFilePath);
                    }
                    File.Move(tempFilePath, fullFilePath);
                    Log.Information("Excel file successfully downloaded to {FullFilePath}", fullFilePath);
                    return true;
                }
                catch (IOException ex) when (attempt < maxRetries)
                {
                    Log.Warning(ex, "File access failed on attempt {Attempt} for {TempFilePath}. Retrying in 1 second.", attempt, tempFilePath);
                    await Task.Delay(1000);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "Failed to download Excel file to {FullFilePath}", fullFilePath);
                    throw;
                }
                finally
                {
                    if (File.Exists(tempFilePath))
                    {
                        try { File.Delete(tempFilePath); } catch { }
                    }
                }
            }
            Log.Error("Failed to download Excel file after {MaxRetries} attempts", maxRetries);
            return false;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Unexpected error downloading Excel file");
            return false;
        }
        finally
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    public static async Task SaveEmployeeDictionaryAsync(string filePath)
    {
        try
        {
            var options = new JsonSerializerOptions { WriteIndented = true };
            string json;
            lock (_dictLock)
            {
                Log.Debug("Serializing EmployeeDict to JSON");
                json = JsonSerializer.Serialize(App.EmployeeDict, options);
            }
            await File.WriteAllTextAsync(filePath, json);
            Log.Information("Employee Dictionary Saved: {FilePath}", filePath);
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Employee Dictionary NOT Saved: {FilePath}", filePath);
        }
    }

    public static string ExportShiftLogWithSignInStatus(Dictionary<int, Employee> employeeDict, string shiftType)
    {
        const string singleline = "+---------------------------------------------------------------------+";
        const string generaltitle = "| Name                         | ID       | Sign In     | Sign Out    |";
        const string signedintitle = "|                              Signed In                              |";
        const string notsignedintitle = "|                            Not Signed In                            |";
        const string tabledivider = "+------------------------------+----------+-------------+-------------+";

        var csvBuilder = new StringBuilder();
        lock (_dictLock)
        {
            csvBuilder.AppendLine(singleline);
            csvBuilder.AppendLine(signedintitle);
            csvBuilder.AppendLine(tabledivider);
            csvBuilder.AppendLine(generaltitle);
            csvBuilder.AppendLine(tabledivider);

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

            foreach (var employee in employeeDict.Values)
            {
                if (employee.ShiftType == shiftType && !employee.SignInTime.HasValue)
                {
                    csvBuilder.AppendLine(employee.ToAsciiTableRow());
                }
            }

            csvBuilder.AppendLine(tabledivider);
        }
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
                    new { type = "TableCell", items = new object[] { new { type = "TextBlock", weight = "Bolder", text = "Name" } } },
                    new { type = "TableCell", items = new object[] { new { type = "TextBlock", weight = "Bolder", text = "ID" } } },
                    new { type = "TableCell", items = new object[] { new { type = "TextBlock", weight = "Bolder", text = "Sign In" } } },
                    new { type = "TableCell", items = new object[] { new { type = "TextBlock", weight = "Bolder", text = "Sign Out" } } }
                }
            }
        };

        lock (_dictLock)
        {
            foreach (var employee in employeeDict.Values)
            {
                if (employee.SignInTime.HasValue == signedIn && employee.ShiftType == shiftType)
                {
                    rows.Add(new
                    {
                        type = "TableRow",
                        cells = new object[]
                        {
                            new { type = "TableCell", items = new object[] { new { type = "TextBlock", text = employee.Name } } },
                            new { type = "TableCell", items = new object[] { new { type = "TextBlock", text = employee.EmployeeNumber.ToString() } } },
                            new { type = "TableCell", items = new object[] { new { type = "TextBlock", text = employee.FormattedSignInTime } } },
                            new { type = "TableCell", items = new object[] { new { type = "TextBlock", text = employee.FormattedSignOutTime ?? "No Sign Out" } } }
                        }
                    });
                }
            }
        }
        return rows.ToArray();
    }

    public static async Task<bool> SendEmployeeDataAsync(string email, string shiftType)
    {
        var requestId = Guid.NewGuid().ToString();
        Log.Information("Sending employee data to {Email} for shift {ShiftType} [RequestId: {RequestId}]", email, shiftType, requestId);
        try
        {
            var jsonData = PrepareJsonForEmail(email, shiftType, requestId);
            using var client = new HttpClient();
            var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

            const int maxRetries = 3;
            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    Log.Debug("Posting to {EmailTableUrl} (Attempt {Attempt}) [RequestId: {RequestId}]", EmailTableUrl, attempt, requestId);
                    var response = await client.PostAsync(EmailTableUrl, content);
                    response.EnsureSuccessStatusCode();
                    Log.Information("Successfully sent employee data to {Email} for shift {ShiftType} [RequestId: {RequestId}]", email, shiftType, requestId);
                    return true;
                }
                catch (HttpRequestException ex) when (attempt < maxRetries)
                {
                    Log.Warning(ex, "HTTP request failed on attempt {Attempt} [RequestId: {RequestId}]. Retrying in 1 second.", attempt, requestId);
                    await Task.Delay(1000);
                }
            }
            Log.Error("Failed to send employee data after {MaxRetries} attempts [RequestId: {RequestId}]", maxRetries, requestId);
            return false;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Unexpected error sending employee data [RequestId: {RequestId}]", requestId);
            return false;
        }
    }

    private static string PrepareJsonForEmail(string email, string shiftType, string requestId)
    {
        lock (_dictLock)
        {
            var signedInList = GetEmployeeDataForJson(shiftType, true);
            var notSignedInList = GetEmployeeDataForJson(shiftType, false);
            var message = new
            {
                email,
                signedIn = signedInList,
                notSignedIn = notSignedInList,
                requestId
            };
            return JsonSerializer.Serialize(message);
        }
    }

    private static object[] GetEmployeeDataForJson(string shiftType, bool signedIn = true)
    {
        var employeeList = new List<object>();
        lock (_dictLock)
        {
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
        }
        Log.Debug("Prepared {Count} {SignedIn} employees for shift {ShiftType}", employeeList.Count, signedIn ? "signed-in" : "not signed-in", shiftType);
        return employeeList.ToArray();
    }

    public static async Task<bool> ExportAdaptiveCardFromTemplateAsync(Dictionary<int, Employee> employeeDict, string shiftType, string message = "")
    {
        var requestId = Guid.NewGuid().ToString();
        Log.Information("Exporting adaptive card for shift {ShiftType} [RequestId: {RequestId}]", shiftType, requestId);
        try
        {
            var shiftTitle = shiftType == "DS" ? "Day Shift" : shiftType == "NS" ? "Night Shift" : "Unknown Shift";
            object teamsMessage;

            // Synchronize only the critical section accessing employeeDict
            lock (_dictLock)
            {
                teamsMessage = new
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
                                new { type = "TextBlock", size = "Medium", weight = "Bolder", text = $"Attendance Report - {shiftTitle} [RequestId: {requestId}]" },
                                new { type = "TextBlock", text = "Signed In", wrap = true },
                                new { type = "Table", columns = new object[] { new { width = 2 }, new { width = 1 }, new { width = 1 }, new { width = 1 } }, rows = GetEmployeeRows(employeeDict, shiftType) },
                                new { type = "TextBlock", text = "Not Signed In", wrap = true },
                                new { type = "Table", columns = new object[] { new { width = 2 }, new { width = 1 }, new { width = 1 }, new { width = 1 } }, rows = GetEmployeeRows(employeeDict, shiftType, false) }
                            }
                        }
                    }
                    }
                };
            }

            // Serialize and log outside the lock
            var firstAttachment = ((dynamic)teamsMessage).attachments[0];
            string firstAttachmentJson = JsonSerializer.Serialize(firstAttachment, new JsonSerializerOptions { WriteIndented = true });
            var byteCount = Encoding.UTF8.GetByteCount(firstAttachmentJson);
            Log.Information("Adaptive card JSON is {ByteCount} bytes [RequestId: {RequestId}]", byteCount, requestId);

            // Perform the async HTTP operation outside the lock
            using var client = new HttpClient();
            var json = JsonSerializer.Serialize(teamsMessage, new JsonSerializerOptions { WriteIndented = true });
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            const int maxRetries = 3;
            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    Log.Debug("Posting to {TeamsTableUrl} (Attempt {Attempt}) [RequestId: {RequestId}]", TeamsTableUrl, attempt, requestId);
                    var response = await client.PostAsync(TeamsTableUrl, content);
                    response.EnsureSuccessStatusCode();
                    Log.Information("Successfully posted adaptive card for shift {ShiftType} [RequestId: {RequestId}]", shiftType, requestId);
                    return true;
                }
                catch (HttpRequestException ex) when (attempt < maxRetries)
                {
                    Log.Warning(ex, "HTTP request failed on attempt {Attempt} [RequestId: {RequestId}]. Retrying in 1 second.", attempt, requestId);
                    await Task.Delay(1000);
                }
            }
            Log.Error("Failed to post adaptive card after {MaxRetries} attempts [RequestId: {RequestId}]", maxRetries, requestId);
            return false;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Error posting adaptive card for shift {ShiftType} [RequestId: {RequestId}]", shiftType, requestId);
            return false;
        }
    }

    //public static async Task<bool> ExportAdaptiveCardFromTemplateAsync(Dictionary<int, Employee> employeeDict, string shiftType, string message = "")
    //{
    //    var requestId = Guid.NewGuid().ToString();
    //    Log.Information("Exporting adaptive card for shift {ShiftType} [RequestId: {RequestId}]", shiftType, requestId);
    //    try
    //    {
    //        var shiftTitle = shiftType == "DS" ? "Day Shift" : shiftType == "NS" ? "Night Shift" : "Unknown Shift";
    //        lock (_dictLock)
    //        {
    //            var teamsMessage = new
    //            {
    //                type = "message",
    //                attachments = new object[]
    //                {
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
    //                                new { type = "TextBlock", size = "Medium", weight = "Bolder", text = $"Attendance Report - {shiftTitle} [RequestId: {requestId}]" },
    //                                new { type = "TextBlock", text = "Signed In", wrap = true },
    //                                new { type = "Table", columns = new object[] { new { width = 2 }, new { width = 1 }, new { width = 1 }, new { width = 1 } }, rows = GetEmployeeRows(employeeDict, shiftType) },
    //                                new { type = "TextBlock", text = "Not Signed In", wrap = true },
    //                                new { type = "Table", columns = new object[] { new { width = 2 }, new { width = 1 }, new { width = 1 }, new { width = 1 } }, rows = GetEmployeeRows(employeeDict, shiftType, false) }
    //                            }
    //                        }
    //                    }
    //                }
    //            };

    //            var firstAttachment = ((dynamic)teamsMessage).attachments[0];
    //            string firstAttachmentJson = JsonSerializer.Serialize(firstAttachment, new JsonSerializerOptions { WriteIndented = true });
    //            var byteCount = Encoding.UTF8.GetByteCount(firstAttachmentJson);
    //            Log.Information("Adaptive card JSON is {ByteCount} bytes [RequestId: {RequestId}]", byteCount, requestId);

    //            using var client = new HttpClient();
    //            var json = JsonSerializer.Serialize(teamsMessage, new JsonSerializerOptions { WriteIndented = true });
    //            var content = new StringContent(json, Encoding.UTF8, "application/json");

    //            const int maxRetries = 3;
    //            for (int attempt = 1; attempt <= maxRetries; attempt++)
    //            {
    //                try
    //                {
    //                    Log.Debug("Posting to {TeamsTableUrl} (Attempt {Attempt}) [RequestId: {RequestId}]", TeamsTableUrl, attempt, requestId);
    //                    var response = await client.PostAsync(TeamsTableUrl, content);
    //                    response.EnsureSuccessStatusCode();
    //                    Log.Information("Successfully posted adaptive card for shift {ShiftType} [RequestId: {RequestId}]", shiftType, requestId);
    //                    return true;
    //                }
    //                catch (HttpRequestException ex) when (attempt < maxRetries)
    //                {
    //                    Log.Warning(ex, "HTTP request failed on attempt {Attempt} [RequestId: {RequestId}]. Retrying in 1 second.", attempt, requestId);
    //                    await Task.Delay(1000);
    //                }
    //            }
    //            Log.Error("Failed to post adaptive card after {MaxRetries} attempts [RequestId: {RequestId}]", maxRetries, requestId);
    //            return false;
    //        }
    //    }
    //    catch (Exception ex)
    //    {
    //        Log.Error(ex, "Error posting adaptive card for shift {ShiftType} [RequestId: {RequestId}]", shiftType, requestId);
    //        return false;
    //    }
    //}


    public static async Task<bool> CheckMostRecentReportDate(string urlForWorkforceJobs, string authToken)
    {
        try
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            if (!string.IsNullOrEmpty(authToken))
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authToken);
            }

            Log.Debug("Checking report date at {Url}", urlForWorkforceJobs);
            var response = await client.GetAsync(urlForWorkforceJobs);
            response.EnsureSuccessStatusCode();
            Log.Debug("Response received. Status Code: {StatusCode}", response.StatusCode);

            var responseBody = await response.Content.ReadAsStringAsync();
            using var jsonDocument = JsonDocument.Parse(responseBody);

            if (jsonDocument.RootElement.ValueKind != JsonValueKind.Array)
            {
                Log.Error("Response is not a JSON array.");
                return false;
            }

            var mostRecentReportArray = jsonDocument.RootElement.EnumerateArray().ToArray();
            Log.Information("Found {Count} report(s).", mostRecentReportArray.Length);

            if (mostRecentReportArray.Length == 0)
            {
                Log.Error("No reports found in the response.");
                return false;
            }

            var report = mostRecentReportArray[0];
            if (!report.TryGetProperty("savedReportId", out var savedReportIdElement) || savedReportIdElement.ValueKind != JsonValueKind.String)
            {
                Log.Error("Error: 'savedReportId' property is missing or invalid.");
                return false;
            }

            if (!report.TryGetProperty("eventDate", out var eventDateElement) || eventDateElement.ValueKind != JsonValueKind.String)
            {
                Log.Error("Error: 'eventDate' property is missing or invalid.");
                return false;
            }

            if (!report.TryGetProperty("files", out var filesElement) || filesElement.ValueKind != JsonValueKind.Array)
            {
                Log.Error("Error: 'files' property is missing or not an array.");
                return false;
            }

            var savedReportId = savedReportIdElement.GetString();
            var mostRecentEventDate = eventDateElement.GetString();
            var filesArray = filesElement.EnumerateArray().ToArray();

            Log.Information("Report ID: {ReportId}, Event Date: {EventDate}, Files Count: {FilesCount}", savedReportId, mostRecentEventDate, filesArray.Length);

            if (filesArray.Length == 0)
            {
                Log.Error("Error: 'files' subarray is empty.");
                return false;
            }

            var firstFile = filesArray[0];
            if (!firstFile.TryGetProperty("fileId", out var fileIdElement) || fileIdElement.ValueKind != JsonValueKind.String ||
                !firstFile.TryGetProperty("fileName", out var fileNameElement) || fileNameElement.ValueKind != JsonValueKind.String ||
                !firstFile.TryGetProperty("fileUri", out var fileUriElement) || fileUriElement.ValueKind != JsonValueKind.String)
            {
                Log.Error("Error: First file in 'files' subarray is missing required properties (fileId, fileName, or fileUri).");
                return false;
            }

            WorkforceReportDownloadUri = fileUriElement.GetString();
            Log.Information("Updated WorkforceReportDownloadUri: {Uri}, File: {FileName} (ID: {FileId})", WorkforceReportDownloadUri, fileNameElement.GetString(), fileIdElement.GetString());

            if (!DateTime.TryParse(mostRecentEventDate, out var eventDate))
            {
                Log.Error("Invalid date format for eventDate: {EventDate}", mostRecentEventDate);
                return false;
            }

            var today = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, TimeZoneInfo.FindSystemTimeZoneById("W. Australia Standard Time")).Date;
            if (eventDate.Date == today)
            {
                Log.Information("Match: The most recent report date ({EventDate}) is today in AWST.", eventDate.ToString("yyyy-MM-dd"));
                return true;
            }

            Log.Information("No match: The most recent report date ({EventDate}) is not today in AWST.", eventDate.ToString("yyyy-MM-dd"));
            return false;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to check most recent report date");
            return false;
        }
    }
}