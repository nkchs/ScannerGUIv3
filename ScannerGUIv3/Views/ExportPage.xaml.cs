using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using ScannerGUIv3.Services;
using ScannerGUIv3.ViewModels;
using System;
using System.Threading.Tasks;
using Windows.Storage.Pickers;
using Windows.Storage;
using Microsoft.UI; // For WindowId
using Microsoft.UI.Windowing; // For AppWindow
using WinRT.Interop; // For WindowNative and InitializeWithWindow

namespace ScannerGUIv3.Views;

public sealed partial class ExportPage : Page
{
    private readonly LogImportExportService _logService;

    public ExportViewModel ViewModel
    {
        get;
    }

    public ExportPage()
    {
        ViewModel = App.GetService<ExportViewModel>();
        _logService = App.GetService<LogImportExportService>();
        InitializeComponent();
    }

    private async void exportButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedOption = ExportOptionsGroup.SelectedItem as ComboBoxItem;
        if (selectedOption == null)
        {
            await ShowMessage("Error", "Please select an export option.");
            return;
        }

        var shiftLog = LogImportExportService.ExportDayShiftLog(App.EmployeeDict);
        bool success = false;
        string option = selectedOption.Content.ToString();

        switch (option)
        {
            case "Email":
                var emailAddress = EmailAddressTextBox.Text;
                if (string.IsNullOrWhiteSpace(emailAddress))
                {
                    await ShowMessage("Error", "Please enter an email address.");
                    return;
                }
                success = await LogImportExportService.ExportToWeb(LogImportExportService.emailUrl,
                    new
                    {
                        email = emailAddress,
                        message = shiftLog
                    });
                EmailAddressTextBox.Text = "";
                break;

            case "Teams":
                success = await LogImportExportService.ExportToWeb(LogImportExportService.teamsUrl,
                    new
                    {
                        message = shiftLog
                    });
                break;

            case "File":
                success = await SaveToFile(shiftLog);
                break;

            default:
                await ShowMessage("Error", "Invalid export option.");
                return;
        }

        if (success)
        {
            await ShowMessage("Success", $"Exported successfully via {option}.");
        }
        else
        {
            await ShowMessage("Error", $"Failed to export via {option}.");
        }
    }

    private async Task<bool> SaveToFile(string content)
    {
        try
        {
            var savePicker = new FileSavePicker();
            savePicker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            savePicker.FileTypeChoices.Add("CSV File", new List<string> { ".csv" });
            savePicker.SuggestedFileName = "ShiftLog";

            // Get the current window handle and initialize the picker
            var hwnd = WindowNative.GetWindowHandle(App.MainWindow); // Assumes App.MainWindow is your main window
            InitializeWithWindow.Initialize(savePicker, hwnd);

            var file = await savePicker.PickSaveFileAsync();
            if (file != null)
            {
                await FileIO.WriteTextAsync(file, content);
                return true;
            }
            return false; // User cancelled
        }
        catch (Exception ex)
        {
            await ShowMessage("Error", $"Failed to save file: {ex.Message}");
            return false;
        }
    }


    private async Task ShowMessage(string title, string message)
    {
        var dialog = new ContentDialog
        {
            Title = title,
            Content = message,
            CloseButtonText = "OK",
            XamlRoot = this.XamlRoot // Set the XamlRoot to the current page's XamlRoot
        };
        await dialog.ShowAsync();
    }

    private async void testButton_Click(object sender, RoutedEventArgs e)
    {
        bool success = await LogImportExportService.DownloadExcelFileAsync("C:/Users/Public/Documents", "Roster");
        if (success)
        {
            await ShowMessage("Success", "Roster downloaded and saved successfully.");
        }
        else
        {
            await ShowMessage("Error", "Failed to download and save roster.");
        }
    }
}
