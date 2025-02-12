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
    public IHost Host
    {
        get;
    }

    public static T GetService<T>()
        where T : class
    {
        if ((App.Current as App)!.Host.Services.GetService(typeof(T)) is not T service)
        {
            throw new ArgumentException($"{typeof(T)} needs to be registered in ConfigureServices within App.xaml.cs.");
        }

        return service;
    }

    public static WindowEx MainWindow { get; set; } = new MainWindow();

    public static UIElement? AppTitlebar
    {
        get; set;
    }


    // ########## Variable Declarations ########## //
    // Time variables //
    public static DateTime currentDate = DateTime.Now;
    public static Calendar calendar = CultureInfo.CurrentCulture.Calendar;
    //public static int weekNumber = calendar.GetWeekOfYear(currentDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
    
    public static DateTime today = DateTime.Today;
    // Define start and end times for day and night shifts
    public static DateTime dayShiftStart = today.AddHours(6);   // 6am on the same day
    public static DateTime dayShiftEnd = today.AddHours(18);    // 6pm on the same day

    public static DateTime nightShiftStart = today.AddHours(18); // 6pm on the same day
    public static DateTime nightShiftEnd = today.AddDays(1).AddHours(6); // 6am on the following day

    // URLs and Strings //
    //public static string sharepointBaseURL = @"https://newcrestmining.sharepoint.com/:f:/r/teams/
    //                                           TelferMaint-Mill/Shared%20Documents/Attendance%20Register/
    //                                           FPM%20Daily%20Sign%20On/Development";
    //public static string excelWeeklyAddress = sharepointBaseURL + @"/Week " + weekNumber + ".xlsm";
    //public static string excelWeeklyAddress = @"C:\Users\ChaseN" + @"\Week " + weekNumber + ".xlsx";
    //public static string excelResourcesOnSiteAddress = @"" + "ResourceOnSite_" + currentDate.ToString("yyyyMMdd") + ".xlsx";
    // ########## End Variable Declarations ########## //

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


    // ########## Dictionary Declarations ########## //
    public static Dictionary<int, Employee> EmployeeDict { get; } = new Dictionary<int, Employee>();
    // ########## Dictionary Declarations ########## //

#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
    public App()
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
    {
        // DICTIONARY STUFF
        Dictionary<int, Employee> employeeDict = new Dictionary<int, Employee>();
        // END DICTIONARY STUFF

        // CONFIG Setup START
        var builder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
        Configuration = builder.Build();
        // CONFIG Setup END

        //Console.WriteLine("Initializing App.");      
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
            services.AddTransient<BlankViewModel>();
            services.AddTransient<BlankPage>();
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
        //var resourcesOnSiteExcel = @"https://newcrestmining.sharepoint.com/:x:/r/teams/TelferMaint-Mill/Shared%20Documents/Attendance%20Register/FPM%20Daily%20Sign%20On/Development/ResourceOnSite_20241030043004.xlsx";
        var resourcesOnSiteExcel = @"C:\Users\ChaseN\source\ScannerGUIRepair\Resources\SRF175 Roster to Excel Today_90days.xlsx";
        var resourcesMasterExcel = @"C:\Users\ChaseN\source\ScannerGUIRepair\Resources\SRF195 Profile Master.xlsx";
        _ = InitializeEmployeeDictionaryAsync(EmployeeDict, resourcesOnSiteExcel);

        // TIMER Setup. Enable the daily scheduler. This is the basis of the timers.
        SetupDailyScheduler();
        // END TIMER Setup

        UnhandledException += App_UnhandledException; // From the default generator.
    }

    // Async function to populate the employee dictionary while 
    private async Task InitializeEmployeeDictionaryAsync(Dictionary<int, Employee> employeeDict, string resourcesOnSiteExcel)
    {
        //await Task.Run(() => PopulateEmployeeDictionary(employeeDict, resourcesOnSiteExcel));
        await Task.Run(() => PopulateEmployeeDictionaryUsingXML(employeeDict, resourcesOnSiteExcel));
    }

    ////// ########## Dictionary FUNCS ########## //

    public void PopulateEmployeeDictionaryUsingXML(Dictionary<int, Employee> employeeDict, string excelPath)
    {
        using (SpreadsheetDocument doc = SpreadsheetDocument.Open(excelPath, false))
        {
            WorkbookPart workbookPart = doc.WorkbookPart;

            // Find the sheet named "Report"
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == "Report");

            if (sheet == null)
            {
                Console.WriteLine("Sheet 'Report' not found.");
                return;
            }

            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
            if (sheetData == null) return;

            // Get Roster Start Date from D1 using GetRosterStartDate function
            DateTime rosterStartDate = GetRosterStartDate(worksheetPart);
            Console.WriteLine("Roster Start Date: " + (rosterStartDate != DateTime.MinValue ? rosterStartDate.ToString("dd/MM/yyyy") : "Invalid Date"));

            Dictionary<int, DateTime> dateHeaders = new Dictionary<int, DateTime>();

            // Read Date Headers from D9 to CP9
            Row headerRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == 9);
            if (headerRow != null)
            {
                foreach (Cell cell in headerRow.Elements<Cell>())
                {
                    string cellValue = GetCellValue(cell, workbookPart);
                    if (DateTime.TryParse(cellValue, out DateTime date))
                    {
                        dateHeaders[cell.CellReference.Value[0] - 'D'] = date;
                    }
                }
            }

            // Read employee data from A9 onwards
            foreach (Row row in sheetData.Elements<Row>().Where(r => r.RowIndex >= 10))
            {
                string firstName = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(0), workbookPart);
                string surname = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(1), workbookPart);
                string personnelCodeStr = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(2), workbookPart);

                if (int.TryParse(personnelCodeStr, out int personnelCode))
                {
                    Employee employee = new Employee(personnelCode, firstName + " " + surname, "");


                    foreach (var entry in dateHeaders)
                    {
                        string shiftValue = GetCellValue(row.Elements<Cell>().ElementAtOrDefault(entry.Key + 3), workbookPart);
                        employee.ShiftType = shiftValue;
                    }

                    employeeDict[personnelCode] = employee;
                }
            }
        }
    }


    private static DateTime GetRosterStartDate(WorksheetPart worksheetPart)
        {
            // Get the cell D1
            Cell cell = worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == "D1");

            if (cell == null || cell.CellValue == null)
            {
                Console.WriteLine("D1 is empty or not found.");
                return DateTime.MinValue;
            }

            string rawValue = cell.CellValue.InnerText;
            //Console.WriteLine("D1 Raw Value: " + rawValue);

            if (double.TryParse(rawValue, out double oaDate))
            {
                return DateTime.FromOADate(oaDate);
            }
            else if (DateTime.TryParseExact(rawValue, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
            {
                return parsedDate;
            }

            Console.WriteLine("D1 could not be converted to a valid date.");
            return DateTime.MinValue;
        }


    private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
    {
        if (cell == null || cell.CellValue == null) return string.Empty;
        string value = cell.CellValue.InnerText;
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText;
        }
        return value;
    }
    // ########## Dictionary ########## //


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