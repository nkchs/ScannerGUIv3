using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using ScannerGUIv3.Services;
using ScannerGUIv3.ViewModels;
using Windows.Storage;
using Windows.Storage.Pickers;
using WinRT.Interop; // For WindowNative and InitializeWithWindow

namespace ScannerGUIv3.Views;

public sealed partial class ExportPage : Page
{
    public ExportViewModel ViewModel
    {
        get;
    }

    public ExportPage()
    {
        ViewModel = App.GetService<ExportViewModel>();
        InitializeComponent();
    }

    private async void exportButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (!TryGetSelectedOptions(out var exportOption, out var shiftType))
            {
                return;
            }

            var shiftLog = GetShiftLog(exportOption, shiftType);

            var success = await ProcessExport(exportOption, shiftLog);

            await ShowMessage(
                success ? "Success" : "Error",
                success ? $"Exported successfully via {exportOption}." : $"Failed to export via {exportOption}.");
        }
        catch (Exception ex)
        {
            await ShowMessage("Error", $"An unexpected error occurred: {ex.Message}");
        }
    }

    private bool TryGetSelectedOptions(out string exportOption, out string shiftType)
    {
        exportOption = (ExportOptionsGroup.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty;
        shiftType = (ShiftOptionsGroup.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty;

        if (string.IsNullOrEmpty(exportOption))
        {
            ShowMessage("Error", "Please select an export option.").GetAwaiter().GetResult();
            return false;
        }

        if (!string.IsNullOrEmpty(shiftType))
        {
            return true;
        }

        ShowMessage("Error", "Please select a shift option.").GetAwaiter().GetResult();
        return false;

    }

    private string GetShiftLog(string exportOption, string shiftType)
    {
        var shiftLog = shiftType.Equals("Night", StringComparison.OrdinalIgnoreCase) ? "NS" : "DS";

        if (!exportOption.Equals("Teams", StringComparison.OrdinalIgnoreCase))
        {
            shiftLog = LogImportExportService.ExportShiftLogWithSignInStatus(App.EmployeeDict, shiftType);
        }

        return shiftLog;
    }

    private async Task<bool> ProcessExport(string exportOption, string shiftLog)
    {
        switch (exportOption.ToLowerInvariant())
        {
            case "email":
                if (EmailAddressTextBox == null)
                {
                    await ShowMessage("Error", "Email input control not found.");
                    return false;
                }

                var emailAddress = EmailAddressTextBox.Text?.Trim();
                if (string.IsNullOrWhiteSpace(emailAddress))
                {
                    await ShowMessage("Error", "Please enter an email address.");
                    return false;
                }

                var success = await LogImportExportService.ExportToWeb(
                    LogImportExportService.EmailUrl,
                    new
                    {
                        email = emailAddress,
                        message = shiftLog
                    });

                if (success)
                {
                    EmailAddressTextBox.Text = string.Empty;
                }
                return success;

            case "teams":
                return await LogImportExportService.ExportAdaptiveCardFromTemplateAsync(App.EmployeeDict, shiftLog);

            case "file":
                return await SaveToFile(shiftLog);

            default:
                await ShowMessage("Error", "Invalid export option.");
                return false;
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
        var success = await LogImportExportService.DownloadExcelFileAsync("C:/Users/Public/Documents", "Roster");
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

// Retired
//private async void exportButton_Click(object sender, RoutedEventArgs e)
//{
//    try
//    {
//        //var shiftType = "DS"; // TODO capture from UI
//        var shiftLog = "DS";
//        bool success;

//        if (ExportOptionsGroup?.SelectedItem is not ComboBoxItem selectedOption)
//        {
//            await ShowMessage("Error", "Please select an export option.");
//            return;
//        }

//        if (ShiftOptionsGroup?.SelectedItem is not ComboBoxItem selectedShift)
//        {
//            await ShowMessage("Error", "Please select a shift option.");
//            return;
//        }

//        var option = selectedOption.Content?.ToString() ?? "";
//        var shiftType = selectedShift.Content?.ToString();

//        if (string.IsNullOrEmpty(option))
//        {
//            await ShowMessage("Error", "Invalid export option selected.");
//            return;
//        }

//        if (selectedShift.Content?.ToString() == "Night")
//        {
//            shiftLog = "NS";
//        }

//        if (!string.Equals(option, "Teams", StringComparison.OrdinalIgnoreCase))
//        {
//            shiftLog = LogImportExportService.ExportShiftLogWithSignInStatus(App.EmployeeDict, shiftType);
//        }

//        switch (option.ToLowerInvariant())
//        {
//            case "email":
//                if (EmailAddressTextBox == null)
//                {
//                    await ShowMessage("Error", "Email input control not found.");
//                    return;
//                }
//                var emailAddress = EmailAddressTextBox.Text;
//                if (string.IsNullOrWhiteSpace(emailAddress))
//                {
//                    await ShowMessage("Error", "Please enter an email address.");
//                    return;
//                }

//                success = await LogImportExportService.ExportToWeb(LogImportExportService.EmailUrl,
//                    new
//                    {
//                        email = emailAddress,
//                        message = shiftLog
//                    });
//                EmailAddressTextBox.Text = "";
//                break;

//            case "teams":
//                success = await LogImportExportService.ExportAdaptiveCardFromTemplateAsync(App.EmployeeDict, shiftType);
//                break;

//            case "file":
//                success = await SaveToFile(shiftLog);
//                break;

//            default:
//                await ShowMessage("Error", "Invalid export option.");
//                return;
//        }

//        await ShowMessage(success ? "Success" : "Error",
//            success ? $"Exported successfully via {option}." : $"Failed to export via {option}.");
//    }
//    catch (Exception ex)
//    {
//        await ShowMessage("Error", $"An unexpected error occurred: {ex.Message}");
//    }
//}