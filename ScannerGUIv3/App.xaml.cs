using System.Globalization;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.UI.Xaml;
using Timer = System.Timers.Timer;
using Microsoft.UI.Dispatching;
using ScannerGUIv3.Activation;
using ScannerGUIv3.Contracts.Services;
using ScannerGUIv3.Core.Contracts.Services;
using ScannerGUIv3.Core.Services;
using ScannerGUIv3.Models;
using ScannerGUIv3.Services;
using ScannerGUIv3.ViewModels;
using ScannerGUIv3.Views;
using ScannerGUIv3.Definitions;
using Application = Microsoft.UI.Xaml.Application;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using ScannerGUIv3.Core;

namespace ScannerGUIv3;

// To learn more about WinUI 3, see https://docs.microsoft.com/windows/apps/winui/winui3/
public partial class App : Application
{
    public IConfiguration Configuration
    {
        get;
    }
    // The .NET Generic Host provides dependency injection, configuration, logging, and other services.
    // https://docs.microsoft.com/dotnet/core/extensions/generic-host
    // https://docs.microsoft.com/dotnet/core/extensions/dependency-injection
    // https://docs.microsoft.com/dotnet/core/extensions/configuration
    // https://docs.microsoft.com/dotnet/core/extensions/logging
    public static T GetService<T>()
        where T : class
    {
        if ((App.Current as App)!.Host.Services.GetService(typeof(T)) is not T service)
        {
            throw new ArgumentException($"{typeof(T)} needs to be registered in ConfigureServices within App.xaml.cs.");
        }

        return service;
    }
    public IHost Host
    {
        get;
    }


    public static WindowEx MainWindow { get; set; } = new MainWindow();

    public static UIElement? AppTitlebar
    {
        get; set;
    }

    
    public List<string> personnelCodes = new();
    public static Dictionary<int, Employee> EmployeeDict { get; } = new Dictionary<int, Employee>();
    //AppState.PersonnelCodesLoaded = false;
    //AppState;

    // ########## Start Timer Declarations ########## //
    //public DispatcherTimer _timer;

    public Timer timer;
    public readonly TimeSpan[] scheduleTimes =
    {
        new(15,55,0),
        new(15,56,0),
        new(15,57,0),
        new(15,58,0),
        new(15,59,0),
    };
    // ########## End Timer Declarations ########## //

#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
    
    
    
    public App()
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
    {
        var builder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
        Configuration = builder.Build();    // CONFIG Setup END

        Console.WriteLine("Initializing App.");
        InitializeComponent();

        Host = Microsoft.Extensions.Hosting.Host.
        CreateDefaultBuilder().
        UseContentRoot(AppContext.BaseDirectory).
        ConfigureServices((context, services) =>
        {
            // Default Activation Handler
            services.AddTransient<ActivationHandler<LaunchActivatedEventArgs>, DefaultActivationHandler>();

            // Other Activation Handlers

            // Services
            services.AddSingleton<ILocalSettingsService, LocalSettingsService>();
            services.AddSingleton<IThemeSelectorService, ThemeSelectorService>();
            services.AddTransient<INavigationViewService, NavigationViewService>();

            services.AddSingleton<IActivationService, ActivationService>();
            services.AddSingleton<IPageService, PageService>();
            services.AddSingleton<INavigationService, NavigationService>();

            // Core Services
            //services.AddSingleton<ISampleDataService, SampleDataService>();
            services.AddSingleton<IFileService, FileService>();

            // Views and ViewModels
            services.AddTransient<SettingsViewModel>();
            services.AddTransient<SettingsPage>();
            services.AddTransient<DataGridViewModel>();
            services.AddTransient<DataGridPage>();
            //services.AddTransient<BlankViewModel>();
            //services.AddTransient<BlankPage>();
            services.AddTransient<MainViewModel>();
            services.AddTransient<MainPage>();
            services.AddTransient<ShellPage>();
            services.AddTransient<ShellViewModel>();

            // Configuration
            services.Configure<LocalSettingsOptions>(context.Configuration.GetSection(nameof(LocalSettingsOptions)));
        }).
        Build();

        // TODO: This logic needs to be updated to find the excel.
        // Async function to fill the dictionary with employee values.
        var resourcesOnSiteExcel = @"C:\Users\ChaseN\source\ScannerGUIRepair\Resources\SRF175 Roster to Excel Today_90days.xlsx";
        //var resourcesMasterExcel = @"C:\Users\ChaseN\source\ScannerGUIRepair\Resources\SRF195 Profile Master Trimmed.xlsx";
        var resourcesMasterExcel = @"C:\Users\ChaseN\source\ScannerGUIRepair\Resources\SRF195 Profile Master.xlsx";
        _ = InitializeEmployeeDictionaryAsync(EmployeeDict, resourcesOnSiteExcel);
        _ = InitializeEmployeeCodesAsync(personnelCodes, resourcesMasterExcel);

        SetupDailyScheduler();  // TIMER Setup. Enable the daily scheduler. This is the basis of the timers.

        UnhandledException += App_UnhandledException; // From the default generator.

        AppState.PersonnelCodesLoaded = false;
        AppState.EmployeeDictionaryLoaded = false;
    }

    // Async function to populate the employee dictionary while 
    private async Task InitializeEmployeeDictionaryAsync(Dictionary<int, Employee> employeeDict, string resourcesOnSiteExcel)
    {
        await Task.Run(() => PopulateEmployeeDictionaryUsingXML(employeeDict, resourcesOnSiteExcel));
    }


    private async Task InitializeEmployeeCodesAsync(List<string> personnelCodes, string resourcesMasterExcel)
    {
        await Task.Run(() => PopulateEmployeeCodesUsingXML(personnelCodes, resourcesMasterExcel));
    }

    public static async Task PopulateEmployeeCodesUsingXML(List<string> personnelCodes, string filePath)
    {
        using var doc = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = doc.WorkbookPart;
        var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
        if (sheet == null) return;

        // Get the sheet data from the first sheet
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
        if (sheetData == null) return;

        // Iterate through each row starting from row 10 (index 10)
        foreach (var row in sheetData.Elements<Row>().Where(r => r.RowIndex >= 10))
        {
            // Extract column values: Personnel Code (B), Department (G), and Active Status (O)
            var personnelCode = GetCellValue(row, "B", workbookPart); // Column B
            var department = GetCellValue(row, "G", workbookPart);   // Column G
            var activeStatus = GetCellValue(row, "O", workbookPart); // Column O

            // Check conditions: Department starts with "MT" and Active Status is "Yes"
            if (!string.IsNullOrEmpty(department) && department.StartsWith("MT", StringComparison.OrdinalIgnoreCase)
                && activeStatus.Equals("Yes", StringComparison.OrdinalIgnoreCase))
            {
                //Console.WriteLine(personnelCode);
                personnelCodes.Add(personnelCode); // Add personnel code to the list
            }
        }
        AppState.PersonnelCodesLoaded = true;
    }

    ////// ########## Dictionary FUNCS ########## //

    public void PopulateEmployeeDictionaryUsingXML(Dictionary<int, Employee> employeeDict, string excelPath)
    {
        using var doc = SpreadsheetDocument.Open(excelPath, false);
        var workbookPart = doc.WorkbookPart;
        if (workbookPart == null) return;

        // Find the sheet named "Report"
        var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == "Report");
        if (sheet == null)
        {
            Console.WriteLine("Sheet 'Report' not found.");
            return;
        }

        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
        if (sheetData == null) return;

        // Get Roster Start Date from D1 using GetRosterStartDate function
        var rosterStartDate = GetRosterStartDate(worksheetPart);
        Console.WriteLine("Roster Start Date: " + (rosterStartDate != DateTime.MinValue ? rosterStartDate.ToString("dd/MM/yyyy") : "Invalid Date"));

        var dateHeaders = new Dictionary<int, DateTime>();

        // Store the next 8 days starting from rosterStartDate
        for (var i = 0; i < 8; i++)
        {
            dateHeaders[i] = rosterStartDate.AddDays(i);
        }

        // Read employee data from A9 onwards
        foreach (var row in sheetData.Elements<Row>().Where(r => r.RowIndex >= 10))
        {
            var firstName = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(0), workbookPart);
            var surname = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(1), workbookPart);
            var personnelCodeStr = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(2), workbookPart);

            if (int.TryParse(personnelCodeStr, out var personnelCode))
            {
                var employee = new Employee(personnelCode, firstName + " " + surname);

                // Store shift types for the next 8 days
                for (var i = 0; i < 8; i++)
                {
                    var shiftValue = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(i + 3), workbookPart);
                    employee.ShiftSchedule[dateHeaders[i]] = shiftValue;
                }
                employeeDict[personnelCode] = employee;
            }
        }
        AppState.EmployeeDictionaryLoaded = true;
    }


    private static DateTime GetRosterStartDate(WorksheetPart worksheetPart)
        {
            // Get the cell D1
            var cell = worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == "D1");

            if (cell == null || cell.CellValue == null)
            {
                Console.WriteLine("D1 is empty or not found.");
                return DateTime.MinValue;
            }

            var rawValue = cell.CellValue.InnerText;
            //Console.WriteLine("D1 Raw Value: " + rawValue);

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


    //private static string GetCellValue(Row row, int columnIndex, WorkbookPart workbookPart)
    //{
    //    Cell cell = row.Elements<Cell>().ElementAtOrDefault(columnIndex - 1); // Get the correct cell by index
    //    if (cell == null || cell.CellValue == null) return string.Empty;

    //    string value = cell.CellValue.InnerText;

    //    // If the cell is a shared string, resolve its value
    //    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
    //    {
    //        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
    //        if (sharedStringTable == null) return value;

    //        if (int.TryParse(value, out int index) && index >= 0 && index < sharedStringTable.ChildElements.Count)
    //        {
    //            return sharedStringTable.Elements<SharedStringItem>().ElementAt(index).InnerText;
    //        }
    //    }

    //    return value; // Return the value if not a shared string
    //}


    private static string GetCellValue(Row row, string columnLetter, WorkbookPart workbookPart)
    {
        // Find the cell in the row that matches the given column (e.g., "B10")
        var cell = row.Elements<Cell>().FirstOrDefault(c => GetColumnLetter(c.CellReference) == columnLetter);
        if (cell == null || cell.CellValue == null) return string.Empty;

        var value = cell.CellValue.InnerText;

        // Handle shared string values
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
        return Regex.Match(cellReference, "[A-Za-z]+").Value; // Extracts letters (column) from cell reference
    }


    // ########## TIMER FUNCS ########## //
    public void SetupDailyScheduler()
    {
        ScheduleNextTask();
    }


    private void ScheduleNextTask()
    {
        var now = DateTime.Now;
        var timeUntilNextTask = GetNextScheduledTime(now);
        // Set the timer to trigger at the calculated interval
        timer = new Timer(timeUntilNextTask.TotalMilliseconds);
        timer.Elapsed += (sender, e) =>
        {
            timer.Stop();  // Stop the timer temporarily

            // Check if DispatcherQueue is available
            var dispatcherQueue = DispatcherQueue.GetForCurrentThread();
            if (dispatcherQueue != null)
            {
                dispatcherQueue.TryEnqueue(() => PerformScheduledOperation());
            }
            else
            {
                // Perform operation directly if no DispatcherQueue is available
                PerformScheduledOperation();
            }

            // Reschedule the timer for the next time
            ScheduleNextTask();
        };
        timer.Start();
    }


    private TimeSpan GetNextScheduledTime(DateTime now)
    {
        foreach (var time in scheduleTimes)
        {
            //DateTime next = now.Date + now.TimeOfDay + time;
            var next = now.Date + time;
            if (next > now)
            {
                return next - now;
            }
        }
        // If all times are in the past, the next scheduled time is tomorrow at the first time
        return (now.Date.AddDays(1) + scheduleTimes[0]) - now;
    }


    private static void PerformScheduledOperation()
    {
        // Your task code here, which runs at 5:00 AM, 5:00 PM, and midnight
        ConsoleService.WriteLine("Scheduled operation executed at " + DateTime.Now);
    }


    // ########## Management FUNCS ########## //
    private void App_UnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
    {
        // TODO: Log and handle exceptions as appropriate.
        // https://docs.microsoft.com/windows/windows-app-sdk/api/winrt/microsoft.ui.xaml.application.unhandledexception.
    }

    protected async override void OnLaunched(LaunchActivatedEventArgs args)
    {
        base.OnLaunched(args);

        await App.GetService<IActivationService>().ActivateAsync(args);
    }
}