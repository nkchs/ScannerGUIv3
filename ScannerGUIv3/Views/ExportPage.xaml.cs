using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using ScannerGUIv3.Services;

using ScannerGUIv3.ViewModels;

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


    private void exportButton_Click(object sender, RoutedEventArgs e)
    {
        Console.ForegroundColor = ConsoleColor.Blue;
        Console.WriteLine("\nExport");

        var ShiftLog = LogImportExportService.ExportDayShiftLog(App.EmployeeDict);
        _ = LogImportExportService.ExportToWeb(ShiftLog, LogImportExportService.teamsUrl);
        Console.WriteLine(ShiftLog);
        Console.ResetColor();
        Console.WriteLine("");
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
