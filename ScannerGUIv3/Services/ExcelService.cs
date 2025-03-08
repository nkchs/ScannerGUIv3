using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using Application = Microsoft.UI.Xaml.Application;
using System.Text.RegularExpressions;
using ScannerGUIv3.Definitions;
using ScannerGUIv3.Core;

namespace ScannerGUIv3.Services;

public class ExcelService
{
    private readonly List<string> personnelCodes = ((App)Application.Current).personnelCodes;

    public static async Task InitializeEmployeeCodesAsync(List<string> personnelCodes, string resourcesMasterExcel)
    {
        try
        {
            await Task.Run(async () =>
            {
                using var httpClient = new HttpClient();
                var response = await httpClient.GetStringAsync("https://prod-18.australiasoutheast.logic.azure.com:443/workflows/7d90dc45ba274d86992b23406da6a420/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=fbi68AVQta0kznlTwHvggUqjoSuLZar2MME9iTklucY&action=SEND_MT_CODES");

                if (!string.IsNullOrEmpty(response) && response.Contains("Code") && response.EndsWith("Code_End"))
                {
                    var codes = response.Split(new[] { "Code", "Code_End" }, StringSplitOptions.RemoveEmptyEntries)[1].Trim().Split('\n');
                    foreach (var code in codes)
                    {
                        if (int.TryParse(code, out var personnelCode))
                        {
                            personnelCodes.Add(code.Trim());
                        }
                    }
                    AppState.PersonnelCodesLoaded = true;
                }
                else
                {
                    throw new Exception("Invalid response format");
                }
            });
        }
        catch
        {
            await Task.Run(() => PopulateEmployeeCodesUsingXML(personnelCodes, resourcesMasterExcel));
        }
    }

    public static async Task PopulateEmployeeCodesUsingXML(List<string> personnelCodes, string filePath)
    {
        await Task.Run(() =>
        {
            using var doc = SpreadsheetDocument.Open(filePath, false);
            var workbookPart = doc.WorkbookPart;
            var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
            if (sheet == null) return;

            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
            if (sheetData == null) return;

            foreach (var row in sheetData.Elements<Row>().Where(r => r.RowIndex >= 10))
            {
                var personnelCode = GetCellValue(row, "B", workbookPart);
                var department = GetCellValue(row, "G", workbookPart);
                var activeStatus = GetCellValue(row, "O", workbookPart);

                if (!string.IsNullOrEmpty(department) && department.StartsWith("MT", StringComparison.OrdinalIgnoreCase)
                    && activeStatus.Equals("Yes", StringComparison.OrdinalIgnoreCase))
                {
                    personnelCodes.Add(personnelCode);
                }
            }
            AppState.PersonnelCodesLoaded = true;
        });
    }

    public async Task InitializeEmployeeDictionaryAsync(Dictionary<int, Employee> employeeDict, string resourcesOnSiteExcel)
    {
        Console.WriteLine("Populate Employee Dictionary Start @ " + DateTime.Now.ToString("HH:mm:ss"));
        await Task.Run(() => PopulateEmployeeDictionaryUsingXML(employeeDict, resourcesOnSiteExcel));
        //Console.WriteLine("Calling Trim");
        await Task.Run(() => TrimEmployeeDictionaryAsync(employeeDict, personnelCodes));
    }

    public void PopulateEmployeeDictionaryUsingXML(Dictionary<int, Employee> employeeDict, string excelPath)
    {
        // Open the Excel document for reading
        using var doc = SpreadsheetDocument.Open(excelPath, false);
        var workbookPart = doc.WorkbookPart;
        if (workbookPart == null) return; // Exit if the workbook part is null

        // Find the sheet named "Report"
        var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == "Report");
        if (sheet == null)
        {
            Console.WriteLine("Sheet 'Report' not found.");
            return; // Exit if the sheet is not found
        }

        // Get the worksheet part associated with the sheet
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
        if (sheetData == null) return; // Exit if the sheet data is null

        // Get the roster start date from the worksheet
        var rosterStartDate = GetRosterStartDate(worksheetPart);
        Console.WriteLine("Roster Start Date: " + (rosterStartDate != DateTime.MinValue ? rosterStartDate.ToString("dd/MM/yyyy") : "Invalid Date"));

        // Create a dictionary to hold the date headers for the next 8 days
        var dateHeaders = new Dictionary<int, DateTime>();
        for (var i = 0; i < 8; i++)
        {
            dateHeaders[i] = rosterStartDate.AddDays(i);
        }

        // Iterate through each row starting from row index 10
        foreach (var row in sheetData.Elements<Row>().Where(r => r.RowIndex >= 10))
        {
            // Get the first name, surname, and personnel code from the row
            var firstName = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(0), workbookPart);
            var surname = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(1), workbookPart);
            var personnelCodeStr = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(2), workbookPart);

            // If the personnel code is a valid integer, create an Employee object
            if (int.TryParse(personnelCodeStr, out var personnelCode))
            {
                var employee = new Employee(personnelCode, firstName + " " + surname);

                // Populate the employee's shift schedule for the next 8 days
                for (var i = 0; i < 8; i++)
                {
                    var shiftValue = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(i + 3), workbookPart);

                    if (i == 0)
                    {
                        employee.ShiftType = shiftValue;
                    }
                    employee.ShiftSchedule[dateHeaders[i]] = shiftValue;
                }
                // Add the employee to the dictionary
                employeeDict[personnelCode] = employee;
            }
        }
        AppState.EmployeeDictionaryLoaded = true; // Set the state to indicate the employee dictionary is loaded
    }

    private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
    {
        if (cell == null || cell.CellValue == null) return string.Empty;
        var value = cell.CellValue.InnerText;
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText;
        }
        return value;
    }

    private static string GetCellValue(Row row, string columnLetter, WorkbookPart workbookPart)
    {
        var cell = row.Elements<Cell>().FirstOrDefault(c => GetColumnLetter(c.CellReference) == columnLetter);
        if (cell == null || cell.CellValue == null) return string.Empty;

        var value = cell.CellValue.InnerText;
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
            if (sharedStringTable == null) return value;

            if (int.TryParse(value, out var index) && index >= 0 && index < sharedStringTable.ChildElements.Count)
            {
                return sharedStringTable.Elements<SharedStringItem>().ElementAt(index).InnerText;
            }
        }

        return value;
    }

    private static string GetColumnLetter(string cellReference)
    {
        return Regex.Match(cellReference, "[A-Za-z]+").Value;
    }

    private static DateTime GetRosterStartDate(WorksheetPart worksheetPart)
    {
        var cell = worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == "D1");
        if (cell == null || cell.CellValue == null)
        {
            Console.WriteLine("D1 is empty or not found.");
            return DateTime.MinValue;
        }

        var rawValue = cell.CellValue.InnerText;
        if (double.TryParse(rawValue, out var oaDate))
        {
            return DateTime.FromOADate(oaDate);
        }
        else if (DateTime.TryParseExact(rawValue, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDate))
        {
            return parsedDate;
        }

        Console.WriteLine("D1 could not be converted to a valid date.");
        return DateTime.MinValue;
    }
    
    private async Task TrimEmployeeDictionaryAsync(Dictionary<int, Employee> employeeDict, List<string> personnelCodes)
    {
        //Console.WriteLine("Trim Entered");
        if (AppState.EmployeeDictionaryLoaded && AppState.PersonnelCodesLoaded && !AppState.EmployeeDictionaryTrimmed)
        {
            await Task.Run(() =>
            {
                //Console.WriteLine("Roster Trim Start");
                var personnelCodeSet = new HashSet<string>(personnelCodes);
                var keysToRemove = employeeDict.Keys.Where(key => !personnelCodeSet.Contains(key.ToString())).ToList();
                foreach (var key in keysToRemove)
                {
                    employeeDict.Remove(key);
                }
                AppState.EmployeeDictionaryTrimmed = true;
            });
        }
        Console.WriteLine("Trim End         @ " + DateTime.Now.ToString("HH:mm:ss"));
    }
}