using System.Globalization;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.UI.Xaml;
using System.Timers;
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
using ScannerGUIv3.Helpers;
using ScannerGUIv3.Definitions;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using Application = Microsoft.UI.Xaml.Application;
using System.Runtime.InteropServices;

//using Newtonsoft.Json.Linq;
//using System.Net.Http.Headers;
//using System.Text;
//using System;
//using System.Net.Http;
//using System.Threading.Tasks;




namespace ScannerGUIv3;

// To learn more about WinUI 3, see https://docs.microsoft.com/windows/apps/winui/winui3/.
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

    public static WindowEx MainWindow { get; private set; } = new MainWindow();

    public static UIElement? AppTitlebar
    {
        get; set;
    }


    // ########## Variable Declarations ########## //
    public static DateTime currentDate = DateTime.Now;
    public static Calendar calendar = CultureInfo.CurrentCulture.Calendar;
    public static int weekNumber = calendar.GetWeekOfYear(currentDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

    public static string sharepointBaseURL = @"https://newcrestmining.sharepoint.com/:f:/r/teams/TelferMaint-Mill/Shared%20Documents/Attendance%20Register/FPM%20Daily%20Sign%20On/Development";
    public static string excelWeeklyAddress = sharepointBaseURL + @"/Week " + weekNumber + ".xlsm";

    //public static string excelWeeklyAddress = @"C:\Users\ChaseN" + @"\Week " + weekNumber + ".xlsx";
    public static string excelResourcesOnSiteAddress = @"" + "ResourceOnSite_" + currentDate.ToString("yyyyMMdd") + ".xlsx";

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
    //public Dictionary<int, Employee> employeeDict = new();
    public static Dictionary<int, Employee> EmployeeDict { get; } = new Dictionary<int, Employee>();
    // ########## Dictionary Declarations ########## //

    public App()
    {
        //Console.WriteLine("Current Week Number: " + weekNumber);
        // DICTIONARY STUFF
        //Dictionary<int, Employee> employeeDict = new Dictionary<int, Employee>();
        //var resourcesOnSiteExcelUrl = @"https://newcrestmining-my.sharepoint.com/personal/nic_chase_newcrest_com_au/Documents/Documents/Projects/Scanner/ResourceOnSite_20240620043001.xlsx";
        var resourcesOnSiteExcelUrl = @"https://newcrestmining.sharepoint.com/:x:/r/teams/TelferMaint-Mill/Shared%20Documents/Attendance%20Register/FPM%20Daily%20Sign%20On/Development/ResourceOnSite_20241030043004.xlsx";
        // END DICTIONARY STUFF


        // CONFIG Setup START
        var builder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
        Configuration = builder.Build();
        // CONFIG Setup END
        

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
            services.AddSingleton<ISampleDataService, SampleDataService>();
            services.AddSingleton<IFileService, FileService>();

            // Views and ViewModels
            services.AddTransient<SettingsViewModel>();
            services.AddTransient<SettingsPage>();
            services.AddTransient<DataGridViewModel>();
            services.AddTransient<DataGridPage>();
            //services.AddTransient<ContentGridDetailViewModel>();
            //services.AddTransient<ContentGridDetailPage>();
            //services.AddTransient<ContentGridViewModel>();
            //services.AddTransient<ContentGridPage>();
            //services.AddTransient<ListDetailsViewModel>();
            //services.AddTransient<ListDetailsPage>();
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


        // MORE DICTIONARY STUFF
        PopulateEmployeeDictionary(EmployeeDict, resourcesOnSiteExcelUrl);
        // TIMER Setup
        SetupDailyScheduler();
        // END TIMER


        UnhandledException += App_UnhandledException;
    }


    // ########## Dictionary FUNCS ########## //
    public void PopulateEmployeeDictionary(Dictionary<int, Employee> employeeDict, string excelPath)
    {
        var excel_ = new Microsoft.Office.Interop.Excel.Application
        {
            Visible = false,
            //Visible = true,
        };

        var excelWorkbook = excel_.Workbooks.Open(excelPath, ReadOnly: true);
        var excelWorksheet = (Worksheet)excelWorkbook.Sheets[3];
        var excelRange = excelWorksheet.UsedRange;

        var maxRow = excelRange.Rows.Count;
        var maxCol = excelRange.Columns.Count;
        //Console.WriteLine("Rows: " + maxRow);
        //Console.WriteLine("Columns: " + maxCol);

        //////////////////////////////////////////////////////////////////////////////////
        // Find the cell containing "Name", this is the first row in the "Name" column.
        var foundNameCell = excelRange.Find("Name", Type.Missing,
                XlFindLookIn.xlValues, XlLookAt.xlPart,
                XlSearchOrder.xlByRows, XlSearchDirection.xlNext,
                false, Type.Missing, Type.Missing);
        var nameColumnNumber = foundNameCell.Column;
        //Console.WriteLine("Names are in Column: " + foundNameCell.Column);
        //Console.WriteLine("Headers are in Row: " + foundNameCell.Row);
        //////////////////////////////////////////////////////////////////////////////////


        //////////////////////////////////////////////////////////////////////////////////
        // Find the cell containing "Name", this is the first row in the "Name" column.
        var foundShiftStatusCell = excelRange.Find("Shift Status", Type.Missing,
            XlFindLookIn.xlValues, XlLookAt.xlPart,
            XlSearchOrder.xlByRows, XlSearchDirection.xlNext,
            false, Type.Missing, Type.Missing);
        var shiftStatusColumnNumber = foundShiftStatusCell.Column;
        //Console.WriteLine("Shift Status is in Column: " + foundShiftStatusCell.Column);
        //////////////////////////////////////////////////////////////////////////////////


        //////////////////////////////////////////////////////////////////////////////////
        // Find the cell containing "Person #", this is the first row in the "Person #" column.
        var foundPersonNumberCell = excelRange.Find("Person #", Type.Missing,
                XlFindLookIn.xlValues, XlLookAt.xlPart,
                XlSearchOrder.xlByRows, XlSearchDirection.xlNext,
                false, Type.Missing, Type.Missing);
        var personNumberColumnNumber = foundPersonNumberCell.Column;
        //Console.WriteLine("Numbers are in Column: " + personNumberColumnNumber);
        //////////////////////////////////////////////////////////////////////////////////


        for (var i = 1; i < maxRow; i++)
        {
            //var _personName = reducedNameRange.Cells[i, 1].Value2;
            var _personName = excelRange[i, nameColumnNumber].Value2;
            if (_personName == null)
            {
                //Console.WriteLine("i: " + i + " null");
            }
            else if (_personName == "Name")
            {
                Console.Write("i: " + i + " " + _personName + " ");
                Console.Write(excelRange[i, personNumberColumnNumber].Value2 + " ");
                Console.Write(excelRange[i, shiftStatusColumnNumber].Value2 + Environment.NewLine);
            }
            else
            {
                int _personNumber = int.Parse( excelRange[i, personNumberColumnNumber].Value2 );
                // Could change the above to int.TryParse;
                string _shiftType = excelRange[i, shiftStatusColumnNumber].Value2;

                Console.Write("i: " + i + " " + _personName + " ");
                Console.Write(excelRange[i, personNumberColumnNumber].Value2 + " ");
                Console.Write(excelRange[i, shiftStatusColumnNumber].Value2 + Environment.NewLine);

                employeeDict.Add(_personNumber, new Employee(_personNumber, _personName,_shiftType));
            }
        }


        Console.WriteLine("Debug");
        excelWorkbook.Close(false, null, null);
        excel_.Quit();
        Marshal.ReleaseComObject(excelWorkbook);
        Marshal.ReleaseComObject(excel_);
    }

    // ########## Dictionary ########## //


    // ########## TIMER FUNCS ########## //
    // ################################# //
    public void SetupDailyScheduler()
    {
        ScheduleNextTask();
    }


    private void ScheduleNextTask()
    {
        // 
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
            if (next > now) return next - now;
        }
        // If all times are in the past, the next scheduled time is tomorrow at the first time
        return (now.Date.AddDays(1) + scheduleTimes[0]) - now;
    }


    private void PerformScheduledOperation()
    {
        // Your task code here, which runs at 5:00 AM, 5:00 PM, and midnight
        Console.WriteLine("Scheduled operation executed at " + DateTime.Now);
    }


    // ########## Management FUNCS ########## //
    // ################################# //
    private void App_UnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
    {
        // TODO: Log and handle exceptions as appropriate.
        // https://docs.microsoft.com/windows/windows-app-sdk/api/winrt/microsoft.ui.xaml.application.unhandledexception.
    }

    protected override async void OnLaunched(LaunchActivatedEventArgs args)
    {
        base.OnLaunched(args);

        await App.GetService<IActivationService>().ActivateAsync(args);
    }
}