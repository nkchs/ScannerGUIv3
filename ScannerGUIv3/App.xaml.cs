using Application = Microsoft.UI.Xaml.Application;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.UI.Xaml;
using ScannerGUIv3.Activation;
using ScannerGUIv3.Contracts.Services;
using ScannerGUIv3.Core;
using ScannerGUIv3.Core.Contracts.Services;
using ScannerGUIv3.Core.Services;
using ScannerGUIv3.Definitions;
using ScannerGUIv3.Models;
using ScannerGUIv3.Services;
using ScannerGUIv3.ViewModels;
using ScannerGUIv3.Views;
using Serilog;
using Serilog.Sinks.SystemConsole.Themes;
using FileHandling = ScannerGUIv3.Services.FileHandling;

//using Microsoft.UI.Xaml.Controls;

namespace ScannerGUIv3;

public partial class App : Application
{
    // ====================================================================================================================
    public App()
    {
        // Settings
        AppState.Logging = false;

        // Variables
        AppState.ResourcesExcelFolderPath = @"C:\Users\Public\Documents\Scanner";

        // Logger setup
        var customTheme = new AnsiConsoleTheme(
            new Dictionary<ConsoleThemeStyle, string>
            {
                [ConsoleThemeStyle.LevelVerbose] = "\x1b[96m", // Cyan
                [ConsoleThemeStyle.LevelDebug] = "\x1b[94m", // Blue
                [ConsoleThemeStyle.LevelInformation] = "\x1b[92m", // Green
                [ConsoleThemeStyle.LevelWarning] = "\x1b[93m", // Yellow
                [ConsoleThemeStyle.LevelError] = "\x1b[91m", // Red
                [ConsoleThemeStyle.LevelFatal] = "\x1b[95m" // Magenta
            });
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Verbose()
            .WriteTo.Console(outputTemplate: "{LevelColor}{Timestamp:HH:mm:ss} {Level:u3} {Message}\x1b[0m{NewLine}",
                theme: customTheme)
            .WriteTo
            .File(AppState.LogFolder + @"\log-.txt", rollingInterval: RollingInterval.Day) // TODO Change to Log Folder
            .CreateLogger();

        var builder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
        Configuration = builder.Build(); // CONFIG Setup END

        Log.Information("++++++++++ INITIALIZING APP ++++++++++");
        InitializeComponent();

        Host = Microsoft.Extensions.Hosting.Host.CreateDefaultBuilder().UseContentRoot(AppContext.BaseDirectory)
            .ConfigureServices((context, services) =>
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
                services.AddSingleton<IFileService, Core.Services.FileService>();

                // Views and ViewModels
                services.AddTransient<ExportViewModel>();
                services.AddTransient<ExportPage>();
                services.AddTransient<SettingsViewModel>();
                services.AddTransient<SettingsPage>();
                services.AddTransient<DataGridViewModel>();
                services.AddTransient<DataGridPage>();
                services.AddTransient<MainViewModel>();
                services.AddTransient<MainPage>();
                services.AddTransient<ShellPage>();
                services.AddTransient<ShellViewModel>();

                // File Handling
                services.AddSingleton<LogImportExportService>();
                services.AddSingleton<ExcelService>();
                services.AddSingleton<FileHandling>();

                // Scheduling
                services.AddSingleton<ScheduleService>();

                // Configuration
                services.Configure<LocalSettingsOptions>(
                    context.Configuration.GetSection(nameof(LocalSettingsOptions)));
            }).Build();


        // Services
        StartUpAsync(); // Startup functions
        var scheduleService = GetService<ScheduleService>(); // 

        // Exceptions
        UnhandledException += App_UnhandledException; // From the default generator.
    }
    // ====================================================================================================================

    // Variable Declarations ==============================================================================================
    
    private static readonly List<string> DebugEmployeeNumbers =
    [
        "20898", "24781", "22388", "24410", "24065", "11356", "27065", "5565", "15734", "3264", "23485", "22794", "4870",
        "4092", "20304", "18062", "23650", "22280", "23672", "20136328", "22139", "5793", "24687", "22483", "3413", "27048",
        "24719","90117936", "5478", "18680"
    ];

    // Dictionaries Start =================================================================================================
    public static Dictionary<int, Employee> EmployeeDict { get; set; } = [];
    public static Dictionary<int, Employee> EmployeeCrossoverDict { get; set; } = [];
    public static Dictionary<int, Employee> PrevNightShiftDict { get; set; } = [];
    public static Dictionary<int, Employee> NextNightShiftDict { get; set; } = [];
    // Dictionaries End ===================================================================================================

    // Startup FUNCS ======================================================================================================
    private static async Task StartUpAsync()
    {
        Log.Warning("========== STARTUP FUNCTIONS BEGIN ==========");

        //AppState.MaintenanceCodesLoaded = false;
        AppState.EmployeeDictionaryLoaded = false;
        AppState.EmployeeDictionaryTrimmed = false;
        AppState.EmployeeDictionaryRefreshed = false;

        // Download the roster  
        await Task.Run(() =>
            LogImportExportService.DownloadExcelFileAsync(AppState.ResourcesExcelFolderPath, "Roster"));
        // Populate the dictionary  
        await Task.Run(() => ExcelService.PopulateEmployeeDictionaryUsingXml(AppState.ResourcesOnSiteExcelPath));

        AppState.StartUpFunctionsComplete = true;

        SignInEmployees(DebugEmployeeNumbers);

        Log.Warning("========== STARTUP FUNCTIONS END ==========");
    }

    // Sign in for debugging
    private static void SignInEmployees(List<string> employeeNumbers)
    {
        Log.Information("Sign In DEBUG Employees");
        foreach (var number in employeeNumbers)
        {
            try
            {
                if (int.TryParse(number, out var employeeNumber))
                {
                    if (EmployeeDict.TryGetValue(employeeNumber, out var employee))
                    {
                        employee.SignIn();
                    }
                    else
                    {
                        //Log.Warning($"Employee with number {employeeNumber} not found.");
                    }
                }
                else
                {
                    Log.Warning($"Invalid employee number: {number}");
                }
            }
            catch
            {
                // Pass if the number isn't found
            }
        }
    }

    // SYSTEM FUNCS =======================================================================================================
    public static WindowEx MainWindow { get; set; } = new MainWindow();
    public IConfiguration Configuration
    {
        get;
    }
    private IHost Host
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

    // Management FUNCS ===================================================================================================
    protected async override void OnLaunched(LaunchActivatedEventArgs args)
    {
        base.OnLaunched(args);

        await App.GetService<IActivationService>().ActivateAsync(args);
    }
    private static void App_UnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
    {
        // TODO: Log and handle exceptions as appropriate.
        // https://docs.microsoft.com/windows/windows-app-sdk/api/winrt/microsoft.ui.xaml.application.unhandledexception.
    }
    public static UIElement? AppTitlebar
    {
        get; set;
    }
}
