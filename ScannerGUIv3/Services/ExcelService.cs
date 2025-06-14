using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using ScannerGUIv3.Definitions;
using ScannerGUIv3.Core;
using Serilog;

namespace ScannerGUIv3.Services;

public class ExcelService
{
    // VARIABLES


    // MAIN FUNCTIONS
    
    public static async Task TrimEmployeeDictionaryShiftType()
    {
        await Task.Run(() =>
        {
            Log.Verbose($"Trim Shift Type Start [Length: {App.EmployeeDict.Count}]");

            var keysToRemove = App.EmployeeDict.Where(e => e.Value.ShiftType == "OS").Select(e => e.Key).ToList();
            foreach (var key in keysToRemove)
            {
                App.EmployeeDict.Remove(key);
            }
            Log.Verbose($"Trim Shift Type End [Length: {App.EmployeeDict.Count}]");
        });
    }

    public static async Task NightShiftCrossoverAsync()
    {
        await Task.Run(() =>
        {
            Log.Verbose("Night Shift Crossover Start");
            // Create a list of keys to remove
            var keysToRemove = new List<int>();
            // If the employee shift type == "NS" add the employee to the crossover dictionary
            foreach (var employee in App.EmployeeDict.Values.Where(employee => employee.ShiftType == "NS"))
            {
                // Add the employee to the crossover dictionary
                App.EmployeeCrossoverDict[employee.EmployeeNumber] = employee;
                // Add the employee to the list of keys to remove
                keysToRemove.Add(employee.EmployeeNumber);
            }
            // Remove the employees from the main dictionary
            foreach (var key in keysToRemove)
            {
                App.EmployeeDict.Remove(key);
            }

            Log.Verbose("Night Shift Crossover End");
        });
    }

    // RETIRED
    public static void PopulateEmployeeDictionaryUsingXml(string excelPath, bool trimShiftTypes = true)
    {
        var validShiftTypes = new HashSet<string> { "NS", "DS", "D1", "D2" };
        Log.Information($"Populate Employee Dict Start");
        try
        {
            using var doc = SpreadsheetDocument.Open(excelPath, false);
            var workbookPart = doc.WorkbookPart ?? throw new Exception("Workbook part is null.");

            var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == "Report") ?? throw new Exception("Sheet 'Report' not found.");
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault() ?? throw new Exception("Sheet data is null.");
            var rosterStartDate = GetRosterStartDate(worksheetPart);
            if (rosterStartDate == DateTime.MinValue) throw new Exception("Invalid roster start date.");

            var dateHeaders = new Dictionary<int, DateTime>();
            for (var i = 0; i < 8; i++)
            {
                dateHeaders[i] = rosterStartDate.AddDays(i);
            }

            foreach (var row in sheetData.Elements<Row>().Where(r => r.RowIndex >= 10))
            {
                var maintenanceCodestr = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(3), workbookPart);

                if (string.IsNullOrEmpty(maintenanceCodestr) ||
                    !int.TryParse(maintenanceCodestr, out var personnelCode) )
                {
                    continue;
                }
                var firstName = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(0), workbookPart);
                if (trimShiftTypes)
                {
                    var firstShiftValue = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(4), workbookPart);
                    //Console.WriteLine(firstName + @" " + firstShiftValue);
                    if (!validShiftTypes.Contains(firstShiftValue))
                    {
                        continue;
                    }
                }

                //var firstName = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(0), workbookPart);
                var surname = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(1), workbookPart);
                var employee = new Employee(personnelCode, firstName + " " + surname);

                for (var i = 0; i < 8; i++)
                {
                    var shiftValue = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(i + 4), workbookPart);

                    if (i == 0)
                    {
                        employee.ShiftType = shiftValue;
                    }
                    employee.ShiftSchedule[dateHeaders[i]] = shiftValue;
                }
                App.EmployeeDict[personnelCode] = employee;
            }
            //AppState.EmployeeDictionaryLoaded = true;

        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to populate employee dictionary from local Excel file.");
        }
        Log.Information($"Populated Employee Dict [Length: {App.EmployeeDict.Count}]");
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
        var cell = worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == "E1");
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
    
    public static async Task GeneratePreviousNightShiftAsync()
    {
        await Task.Run(() =>
        {
            Log.Verbose("Generate Current Night Shift Dict Start");
            // Iterate through the employee dictionary and add entries with ShiftType "NS" to the new dictionary
            foreach (var employee in App.EmployeeDict.Values.Where(employee => employee.ShiftType == "NS"))
            {
                App.PrevNightShiftDict[employee.EmployeeNumber] = employee;
            }

            Log.Verbose("Generate Current Night Shift Dict End");
            Log.Information($"Copied Night Shift Entries [Length: {App.PrevNightShiftDict.Count}]");
        });
    }

    public static async Task RemovePreviousNightShiftAsync()
    {
        await Task.Run(() =>
        {
            Log.Verbose("Remove Previous Night Shift Start");
            // Create a list of employees to remove (to avoid modifying dictionary during iteration)
            var employeesToRemove = App.EmployeeDict.Values
                .Where(employee => employee.ShiftType == "NS")
                .Select(employee => employee.EmployeeNumber)
                .ToList();

            // Remove the identified employees
            foreach (var employeeNumber in employeesToRemove)
            {
                App.EmployeeDict.Remove(employeeNumber);
            }

            Log.Verbose("Remove Previous Night Shift End");
            Log.Information($"Removed Night Shift Entries [Count: {employeesToRemove.Count}]");
        });
    }

    public static async Task GenerateNextNightShiftAsync()
    {
        await Task.Run(() =>
        {
            Log.Verbose("Generate Next Night Shift Dict Start");
            // Iterate through the employee dictionary and add entries with ShiftType "NS" to the new dictionary
            foreach (var employee in App.EmployeeDict.Values.Where(employee => employee.ShiftType == "NS"))
            {
                App.NextNightShiftDict[employee.EmployeeNumber] = employee;
            }

            Log.Verbose("Generate Next Night Shift Dict End");
            Log.Information($"Copied Night Shift Entries [Length: {App.NextNightShiftDict.Count}]");
        });
    }

    public static async Task InsertCurrentNightShiftAsync()
    {
        Log.Verbose("Insert Current Night Shift Start");
        // Ensure this runs on a background thread to avoid blocking UI
        await Task.Run(() =>
        {
            // Insert all entries from PrevNightShiftDict into EmployeeDict
            foreach (var employee in App.PrevNightShiftDict)
            {
                // If the key already exists, this will overwrite the existing entry
                App.EmployeeDict[employee.Key] = employee.Value;
            }
        });

        // Sort EmployeeDict by Employee.Name and convert to a list
        var sortedEmployees = App.EmployeeDict.OrderBy(employee => employee.Value.Name).ToList();

        // Optional: Update EmployeeDict with sorted order (though Dictionary doesn't maintain order)
        App.EmployeeDict.Clear();
        foreach (var employee in sortedEmployees)
        {
            App.EmployeeDict[employee.Key] = employee.Value;
        }
        Log.Verbose("Insert Current Night Shift End");
    }

    public static async Task InsertNextNightShiftAsync()
    {
        Log.Verbose("Insert Next Night Shift Start");
        // Ensure this runs on a background thread to avoid blocking UI
        await Task.Run(() =>
        {
            // Insert all entries from PrevNightShiftDict into EmployeeDict
            foreach (var employee in App.NextNightShiftDict)
            {
                // If the key already exists, this will overwrite the existing entry
                App.EmployeeDict[employee.Key] = employee.Value;
            }
        });

        // Sort EmployeeDict by Employee.Name and convert to a list
        var sortedEmployees = App.EmployeeDict.OrderBy(employee => employee.Value.Name).ToList();

        // Optional: Update EmployeeDict with sorted order (though Dictionary doesn't maintain order)
        App.EmployeeDict.Clear();
        foreach (var employee in sortedEmployees)
        {
            App.EmployeeDict[employee.Key] = employee.Value;
        }
        Log.Verbose("Insert Next Night Shift End");
    }

    // Retired

}


