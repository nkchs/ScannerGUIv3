using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text.RegularExpressions;
using ScannerGUIv3.Definitions;
using ScannerGUIv3.Core;
using Serilog;

namespace ScannerGUIv3.Services;

public class ExcelService
{
    // VARIABLES


    // MAIN FUNCTIONS
    public static async Task InitializeEmployeeCodesAsyncHTTP(List<string> personnelCodes, string resourcesMasterExcel)
    {
        Log.Verbose("Employee Code DL Start");

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
                            App.PersonnelCodes.Add(code.Trim());
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
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to download employee codes via HTTP. Falling back to local Excel file.");
            await Task.Run(() => PopulateEmployeeCodesUsingXML(App.PersonnelCodes, resourcesMasterExcel));
        }
        Log.Verbose("Employee Code DL End");
    }

    public static async Task PopulateEmployeeCodesUsingXML(List<string> personnelCodes, string filePath)
    {
        await Task.Run(() =>
        {
            try
            {
                using var doc = SpreadsheetDocument.Open(filePath, false);
                var workbookPart = doc.WorkbookPart;
                var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                if (sheet == null) throw new Exception("Sheet not found in the Excel file.");

                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                if (sheetData == null) throw new Exception("Sheet data not found in the Excel file.");

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
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to populate employee codes from local Excel file.");
                AppState.PersonnelCodesLoaded = false;
            }
        });
    }

    public static async Task TrimEmployeeDictionaryAsync(Dictionary<int, Employee> employeeDict, List<string> personnelCodes)
    {
        if (AppState.EmployeeDictionaryLoaded && AppState.PersonnelCodesLoaded && !AppState.EmployeeDictionaryTrimmed)
        {
            await Task.Run(() =>
            {
                Log.Verbose("Trim Start");

                var personnelCodeSet = new HashSet<string>(App.PersonnelCodes);
                var keysToRemove = employeeDict.Keys.Where(key => !personnelCodeSet.Contains(key.ToString())).ToList();
                foreach (var key in keysToRemove)
                {
                    employeeDict.Remove(key);
                }
                AppState.EmployeeDictionaryTrimmed = true;
            });
        }
        Log.Verbose("Trim End");
    }

    public static async Task TrimEmployeeDictionaryShiftType()
    {
        await Task.Run(() =>
        {
            Log.Verbose("Trim Shift Type Start");

            var keysToRemove = App.EmployeeDict.Where(e => e.Value.ShiftType == "OS").Select(e => e.Key).ToList();
            foreach (var key in keysToRemove)
            {
                App.EmployeeDict.Remove(key);
            }
            Log.Verbose("Trim Shift Type End");
        });
    }

    public static async Task NightShiftCrossoverAsync()
    {
        await Task.Run(() =>
        {
            Log.Verbose("Night Shift Crossover Start");

            var keysToRemove = new List<int>();
            foreach (var employee in App.EmployeeDict.Values.Where(employee => employee.ShiftType == "NS"))
            {
                App.EmployeeCrossoverDict[employee.EmployeeNumber] = employee;
                keysToRemove.Add(employee.EmployeeNumber);
            }

            foreach (var key in keysToRemove)
            {
                App.EmployeeDict.Remove(key);
            }

            Log.Verbose("Night Shift Crossover End");
        });
    }


    // RETIRED
    public static async Task InitializeEmployeeDictionaryAsync(Dictionary<int, Employee> employeeDict, string resourcesOnSiteExcel)
    {
        Log.Verbose("Populate Employee Dictionary ASYNC [Using XML]");
        // Populate the dictionary using XML
        await Task.Run(() => PopulateEmployeeDictionaryUsingXML(employeeDict, resourcesOnSiteExcel));
        // Trim the employee dictionary
        await Task.Run(() => TrimEmployeeDictionaryAsync(employeeDict, App.PersonnelCodes));
    }

    public static void PopulateEmployeeDictionaryUsingXML(Dictionary<int, Employee> employeeDict, string excelPath)
    {
        try
        {
            using var doc = SpreadsheetDocument.Open(excelPath, false);
            var workbookPart = doc.WorkbookPart;
            if (workbookPart == null) throw new Exception("Workbook part is null.");

            var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == "Report");
            if (sheet == null) throw new Exception("Sheet 'Report' not found.");

            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
            if (sheetData == null) throw new Exception("Sheet data is null.");

            var rosterStartDate = GetRosterStartDate(worksheetPart);
            if (rosterStartDate == DateTime.MinValue) throw new Exception("Invalid roster start date.");

            var dateHeaders = new Dictionary<int, DateTime>();
            for (var i = 0; i < 8; i++)
            {
                dateHeaders[i] = rosterStartDate.AddDays(i);
            }

            foreach (var row in sheetData.Elements<Row>().Where(r => r.RowIndex >= 10))
            {
                var firstName = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(0), workbookPart);
                var surname = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(1), workbookPart);
                var personnelCodeStr = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(2), workbookPart);

                if (int.TryParse(personnelCodeStr, out var personnelCode))
                {
                    var employee = new Employee(personnelCode, firstName + " " + surname);

                    for (var i = 0; i < 8; i++)
                    {
                        var shiftValue = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(i + 3), workbookPart);

                        if (i == 0)
                        {
                            employee.ShiftType = shiftValue;
                        }
                        employee.ShiftSchedule[dateHeaders[i]] = shiftValue;
                    }
                    App.EmployeeDict[personnelCode] = employee;
                }
            }
            AppState.EmployeeDictionaryLoaded = true;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to populate employee dictionary from local Excel file.");
        }
    }


    // EXCEL HANDLERS & MINOR FUNCTIONS
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
}
