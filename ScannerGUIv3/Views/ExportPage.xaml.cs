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
        var selectedOption = ExportOptionsGroup.SelectedItem as RadioButton;
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

    private void testButton_Click(object sender, RoutedEventArgs e)
    {
        Console.ForegroundColor = ConsoleColor.Green;
        //Console.WriteLine("Export");
        var selectedOption = ExportOptionsGroup.SelectedItem as RadioButton;
        if (selectedOption != null)
        {
            var ShiftLog = LogImportExportService.ExportDayShiftLog(App.EmployeeDict);
            switch (selectedOption.Content)
            {
                default:
                    switch (selectedOption.Content)
                    {
                        case "Email":
                            Console.WriteLine("Exporting via Email...");
                            var emailAddress = EmailAddressTextBox.Text;
                            Console.WriteLine(emailAddress);
                            // Add logic to export via Email
                            break;
                        case "Teams":
                            Console.WriteLine("Exporting via Teams...");
                            // Add logic to export via Teams
                            _ = LogImportExportService.ExportToWeb(ShiftLog, LogImportExportService.teamsUrl);
                            break;
                        case "File":
                            Console.WriteLine("Exporting to File...");
                            // Add logic to export to File
                            break;
                        default:
                            Console.WriteLine("Unknown export option.");
                            break;
                    }
                    break;
            }
            //Console.WriteLine($"Selected Option: {selectedOption.Content}");
        }
        else
        {
            Console.WriteLine("No option selected.");
        }
        //var ShiftLog = LogImportExportService.ExportDayShiftLog(App.EmployeeDict);
        //_ = LogImportExportService.ExportToWeb(ShiftLog, LogImportExportService.teamsUrl);
        //Console.WriteLine(ShiftLog);
        Console.ResetColor();
        //Console.WriteLine("");
    }
}
