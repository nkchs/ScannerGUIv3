using System.Globalization;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.UI.Xaml;

using ScannerGUIv3.Activation;
using ScannerGUIv3.Contracts.Services;
using ScannerGUIv3.Core.Contracts.Services;
using ScannerGUIv3.Core.Services;
using ScannerGUIv3.Models;
using ScannerGUIv3.Services;
using ScannerGUIv3.ViewModels;
using ScannerGUIv3.Views;

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


    // Variable Declarations

    public static DateTime currentDate = DateTime.Now;
    public static Calendar calendar = CultureInfo.CurrentCulture.Calendar;
    public static int weekNumber = calendar.GetWeekOfYear(currentDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

    public static string excelWeeklyAddress = @"C:\Users\ChaseN" + @"\Week " + weekNumber + ".xlsx";
    public static string excelResourceAddress = @"" + "ResourceOnSite_" + currentDate.ToString("yyyyMMdd") + ".xlsx";

    // End Variable Declarations

    public DispatcherTimer _timer;

    // End Timer Declarations

    public App()
    {
        var builder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
        Configuration = builder.Build();
        
        Console.WriteLine("Initializing App.");
        //Console.WriteLine("Current Week Number: " + weekNumber);

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
            services.AddTransient<ContentGridDetailViewModel>();
            services.AddTransient<ContentGridDetailPage>();
            services.AddTransient<ContentGridViewModel>();
            services.AddTransient<ContentGridPage>();
            services.AddTransient<ListDetailsViewModel>();
            services.AddTransient<ListDetailsPage>();
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

        UnhandledException += App_UnhandledException;
    }

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