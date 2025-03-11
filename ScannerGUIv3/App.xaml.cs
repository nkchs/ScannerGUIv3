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
using Microsoft.UI.Xaml.Controls;

namespace ScannerGUIv3;

public partial class App : Application
{
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


        Log.Information("Initializing App @ " + DateTime.Now.ToString("HH:mm:ss"));
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
                services.Configure<LocalSettingsOptions>(
                    context.Configuration.GetSection(nameof(LocalSettingsOptions)));
            }).Build();


        // Services
        StartUpAsync(); // Startup functions
        var scheduleService = GetService<ScheduleService>(); // 

        // Exceptions
        UnhandledException += App_UnhandledException; // From the default generator.
    }



    public static UIElement? AppTitlebar
    {
        get;
        set;
    }


    // Variable Declarations
    //public static readonly List<string> MaintenanceCodes = [];
    public static readonly List<int> MaintenanceCodes = [];

    public static Dictionary<int, Employee> EmployeeDict
    {
        get;
        set;
    } = new();

    public static Dictionary<int, Employee> EmployeeCrossoverDict
    {
        get;
        set;
    } = new();


    public static WindowEx MainWindow
    {
        get;
        set;
    } = new MainWindow();

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
        Log.Warning("STARTUP FUNCTIONS BEGIN");
        //var mainPage = App.GetService<MainPage>();
        //await mainPage.ShowMessage("Initalizing", "Retrieving Data");

        AppState.MaintenanceCodesLoaded = false;
        AppState.EmployeeDictionaryLoaded = false;
        AppState.EmployeeDictionaryTrimmed = false;
        AppState.EmployeeDictionaryRefreshed = false;

        // Download the roster  
        await Task.Run(() => LogImportExportService.DownloadExcelFileAsync(AppState.ResourcesExcelFolderPath, "Roster"));
        // Download the Maintenance codes  
        await Task.Run(() => ExcelService.InitializeMaintenanceCodesHTTP(AppState.ResourcesMasterExcelPath));
        // Populate the dictionary  
        await Task.Run(() => ExcelService.PopulateEmployeeDictionaryUsingXML(EmployeeDict, AppState.ResourcesOnSiteExcelPath));
        // Trim the dictionary of Maintenance codes  
        await Task.Run(() => ExcelService.TrimEmployeeDictionaryAsync(EmployeeDict, MaintenanceCodes));
        // Trim the dictionary of shift types  
        await Task.Run(ExcelService.TrimEmployeeDictionaryShiftType);

        Log.Warning("STARTUP FUNCTIONS END");
    }

    //public static class MessageBoxHelper
    //{
    //    private static ContentDialog? _dialog;

    //    public static async Task ShowMessageAsync(string title, string content)
    //    {
    //        if (_dialog != null)
    //        {
    //            await CloseMessageAsync();
    //        }

    //        _dialog = new ContentDialog
    //        {
    //            Title = title,
    //            Content = content,
    //            CloseButtonText = "Close",
    //            XamlRoot = App.MainWindow.Content?.XamlRoot // Use App.MainWindow to access the instance  
    //        };

    //        await _dialog.ShowAsync();
    //    }

    //    public static async Task CloseMessageAsync()
    //    {
    //        if (_dialog != null)
    //        {
    //            _dialog.Hide();
    //            _dialog = null;
    //        }
    //    }
    //}
}
