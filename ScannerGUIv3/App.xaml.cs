//using System.Globalization;
//using System.Text.RegularExpressions;

using Application = Microsoft.UI.Xaml.Application;

//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Spreadsheet;

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
//using Microsoft.UI.Dispatching;
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

//using Timer = System.Timers.Timer;

namespace ScannerGUIv3;

// To learn more about WinUI 3, see https://docs.microsoft.com/windows/apps/winui/winui3/
public partial class App : Application
{
    public App()
    {
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.Console()
            .WriteTo.File("logs/log-.txt", rollingInterval: RollingInterval.Day)
            .CreateLogger();

        var builder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
        Configuration = builder.Build();    // CONFIG Setup END


        Log.Information("Initializing App @ " + DateTime.Now.ToString("HH:mm:ss"));
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
            services.AddSingleton<IFileService, FileService>();

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

            // Scheduling
            services.AddSingleton<ScheduleService>();

            // Configuration
            services.Configure<LocalSettingsOptions>(context.Configuration.GetSection(nameof(LocalSettingsOptions)));
        }).
        Build();

        // Variables
        AppState.ResourcesExcelFolderPath = @"C:/Users/Public/Documents";

        // Services
        //ExcelService _excelService = GetService<ExcelService>();
        StartUpAsync(); // Startup functions
        //ScheduleService scheduleService = GetService<ScheduleService>();

        // Exceptions
        UnhandledException += App_UnhandledException; // From the default generator.
    }

    public static UIElement? AppTitlebar
    {
        get; set;
    }

    // Variable Declarations
    public static readonly List<string> PersonnelCodes = [];
    
    public static Dictionary<int, Employee> EmployeeDict { get; } = new Dictionary<int, Employee>();

    public static WindowEx MainWindow { get; set; } = new MainWindow();

    public IConfiguration Configuration
    {
        get;
    }

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
    // ########## Startup FUNCS ########## //

    protected async override void OnLaunched(LaunchActivatedEventArgs args)
    {
        base.OnLaunched(args);

        await App.GetService<IActivationService>().ActivateAsync(args);
    }


    // ########## Management FUNCS ########## //
    private void App_UnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
    {
        // TODO: Log and handle exceptions as appropriate.
        // https://docs.microsoft.com/windows/windows-app-sdk/api/winrt/microsoft.ui.xaml.application.unhandledexception.
    }

    private async void StartUpAsync()
    {
        Console.ForegroundColor = ConsoleColor.DarkRed;
        //Console.WriteLine("STARTUP FUNCTIONS START @ " + DateTime.Now.ToString("HH:mm:ss"));
        Log.Warning("STARTUP FUNCTIONS BEGIN");

        AppState.PersonnelCodesLoaded = false;
        AppState.EmployeeDictionaryLoaded = false;
        AppState.EmployeeDictionaryTrimmed = false;
        AppState.EmployeeDictionaryRefreshed = false;
        Console.ForegroundColor = ConsoleColor.DarkRed;

        // Download the roster
        await Task.Run(() => LogImportExportService.DownloadExcelFileAsync(AppState.ResourcesExcelFolderPath, "Roster"));
        // Download the personnel codes
        await Task.Run(() => ExcelService.InitializeEmployeeCodesAsyncHTTP(PersonnelCodes, AppState.ResourcesMasterExcelPath));
        // Populate the dictionary
        await Task.Run(() => ExcelService.PopulateEmployeeDictionaryUsingXML(EmployeeDict, AppState.ResourcesOnSiteExcelPath));
        // Trim the dictionary
        await Task.Run(() => ExcelService.TrimEmployeeDictionaryAsync(EmployeeDict, PersonnelCodes));

        Log.Warning("STARTUP FUNCTIONS END");
        //Console.WriteLine("STARTUP FUNCTIONS END  @ " + DateTime.Now.ToString("HH:mm:ss"));
    }

}