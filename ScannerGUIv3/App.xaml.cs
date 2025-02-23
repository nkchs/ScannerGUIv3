using System.Globalization;
using System.Text.RegularExpressions;

using Application = Microsoft.UI.Xaml.Application;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.UI.Dispatching;
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

using Timer = System.Timers.Timer;

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

    // Variable Declarations

    public string resourcesOnSiteExcel
    {
        get; set;
    }
        public string resourcesMasterExcel
    {
        get; set;
    }
        public List<string> personnelCodes = [];
    public static Dictionary<int, Employee> EmployeeDict { get; } = new Dictionary<int, Employee>();
  
    
    public App()
    {
        var builder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
        Configuration = builder.Build();    // CONFIG Setup END

        Console.WriteLine("Initializing App @ " + DateTime.Now.ToString("HH:mm:ss"));
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


        AppState.PersonnelCodesLoaded = false;
        AppState.EmployeeDictionaryLoaded = false;
        AppState.EmployeeDictionaryTrimmed = false;
        AppState.EmployeeDictionaryRefreshed = false;


        ExcelService _excelService = GetService<ExcelService>();
        Console.WriteLine("Roster Start     @ " + DateTime.Now.ToString("HH:mm:ss"));
        // TODO: This logic needs to be updated to find the excel. // Async function to fill the dictionary with employee values.
        resourcesOnSiteExcel = @"C:\Users\ChaseN\source\ScannerGUIRepair\Resources\SRF175 Roster to Excel Today_90days.xlsx";
        resourcesMasterExcel = @"C:\Users\ChaseN\source\ScannerGUIRepair\Resources\SRF195 Profile Master.xlsx";
        //var resourcesMasterExcel = @"C:\Users\ChaseN\source\ScannerGUIRepair\Resources\SRF195 Profile Master Trimmed.xlsx";
        _ = ExcelService.InitializeEmployeeCodesAsync(personnelCodes, resourcesMasterExcel);
        _ = _excelService.InitializeEmployeeDictionaryAsync(EmployeeDict, resourcesOnSiteExcel);


        ScheduleService _scheduleService = GetService<ScheduleService>();
        //_scheduleService = new ScheduleService();


        UnhandledException += App_UnhandledException; // From the default generator.
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
