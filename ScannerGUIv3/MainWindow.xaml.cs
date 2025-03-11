using ScannerGUIv3.Helpers;
using Microsoft.UI.Windowing;
using Windows.UI.ViewManagement;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.UI.Xaml;
using ScannerGUIv3.Core;
using ScannerGUIv3.Services;
using Serilog;

namespace ScannerGUIv3;

public sealed partial class MainWindow : WindowEx
{
    private Microsoft.UI.Dispatching.DispatcherQueue dispatcherQueue;

    private UISettings settings;

    public MainWindow()
    {
        //Console.WriteLine("Initializing Main Window.");

        InitializeComponent();

        Width = 1000;
        Height = 1300;

        AppWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets/WindowIcon.ico"));
        Content = null;
        Title = "AppDisplayName".GetLocalized();

        // Theme change code picked from https://github.com/microsoft/WinUI-Gallery/pull/1239
        dispatcherQueue = Microsoft.UI.Dispatching.DispatcherQueue.GetForCurrentThread();
        settings = new UISettings();
        settings.ColorValuesChanged += Settings_ColorValuesChanged; // cannot use FrameworkElement.ActualThemeChanged event

        //var windowWidth = Bounds.Width;
        //var windowHeight = Bounds.Height;
        //Console.WriteLine("Window Width: " + windowWidth);
        //Console.WriteLine("Window Height: " + windowHeight);

        // Add the Closed event handler
        Closed += MainWindow_Closed;
    }

    // this handles updating the caption button colors correctly when windows system theme is changed
    // while the app is open
    private void Settings_ColorValuesChanged(UISettings sender, object args)
    {
        // This calls comes off-thread, hence we will need to dispatch it to current app's thread
        dispatcherQueue.TryEnqueue(TitleBarHelper.ApplySystemThemeToCaptionButtons);
    }

    private static async void MainWindow_Closed(object sender, WindowEventArgs e)
    {
        try
        {
            await LogImportExportService.SaveEmployeeDictionaryAsync(AppState.StateFolder + @"/EmployeeDict.json");
            Log.Information("MainWindow is closing.");
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Exception Raised On Exit");
            //throw; // TODO handle exception
        }
    }
}