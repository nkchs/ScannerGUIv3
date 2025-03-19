using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

using ScannerGUIv3.Activation;
using ScannerGUIv3.Contracts.Services;
using ScannerGUIv3.Views;
using Serilog;

namespace ScannerGUIv3.Services;

public class ActivationService(ActivationHandler<LaunchActivatedEventArgs> defaultHandler, IEnumerable<IActivationHandler> activationHandlers, IThemeSelectorService themeSelectorService) : IActivationService
{
    private readonly ActivationHandler<LaunchActivatedEventArgs> _defaultHandler = defaultHandler;
    private readonly IEnumerable<IActivationHandler> _activationHandlers = activationHandlers;
    private readonly IThemeSelectorService _themeSelectorService = themeSelectorService;
    private UIElement? _shell = null;

    public async Task ActivateAsync(object activationArgs)
    {
        //Log.Information("ActivateAsync() from ActivationService.cs");
        //if (App.MainWindow == null)
        //{
            //Log.Information("MainWindow == null | from ActivationService.cs");
        App.MainWindow = new MainWindow();
        App.MainWindow.Activate(); // Check if Activation needs to occur here.
        //}

        // Execute tasks before activation.
        await InitializeAsync();

        // Set the MainWindow Content.
        if (App.MainWindow.Content == null)
        {
            //Console.WriteLine("App.MainWindow.Content == null.");
            _shell = App.GetService<ShellPage>();
            App.MainWindow.Content = _shell ?? new Frame();
        }

        // Handle activation via ActivationHandlers.
        await HandleActivationAsync(activationArgs);

        // Activate the MainWindow.
        App.MainWindow.Activate(); // Check if activation needs to occur here.

        //App.MainWindow.MoveAndResize(App.MainWindow.Bounds.X, App.MainWindow.Bounds.Y, 700, 700);
        //App.MainWindow.MoveAndResize(2000, 700, 700, 700);

        // Execute tasks after activation.
        await StartupAsync();
    }

    private async Task HandleActivationAsync(object activationArgs)
    {
        var activationHandler = _activationHandlers.FirstOrDefault(h => h.CanHandle(activationArgs));

        if (activationHandler != null)
        {
            await activationHandler.HandleAsync(activationArgs);
        }

        if (_defaultHandler.CanHandle(activationArgs))
        {
            await _defaultHandler.HandleAsync(activationArgs);
        }
    }

    private async Task InitializeAsync()
    {
        await _themeSelectorService.InitializeAsync().ConfigureAwait(false);
        await Task.CompletedTask;
    }

    private async Task StartupAsync()
    {
        await _themeSelectorService.SetRequestedThemeAsync();
        await Task.CompletedTask;
    }
}