using Microsoft.UI.Xaml;
using ScannerGUIv3.ViewModels;
using Microsoft.UI.Xaml.Input;
using Windows.System;
using ScannerGUIv3.Services;
using ScannerGUIv3.Core;
using Microsoft.UI.Xaml.Controls;
using Serilog;
using FileHandling = ScannerGUIv3.Services.FileHandling;

namespace ScannerGUIv3.Views;

public sealed partial class MainPage : Microsoft.UI.Xaml.Controls.Page
{

    public MainViewModel ViewModel { get; }

    public MainPage()// 
    {
        ViewModel = App.GetService<MainViewModel>();
        InitializeComponent();
        ConsoleService.Initialize(ConsoleOutput); // Initialize with the console TextBox
        Loaded += OnLoadedAsync;
    }

    private void OnLoadedAsync(object sender, RoutedEventArgs e)
    {
        try
        {
            PersonnelNumberTextBox.Focus(FocusState.Programmatic);
            ShowMessage("Initializing Data", "Please Wait...");
        }
        catch (Exception ex)
        {
            throw; // TODO handle exception
        }
    }

    private void signInButton_Click(object sender, RoutedEventArgs e)
    {
        if (int.TryParse(PersonnelNumberTextBox.Text, out var personnelCode))
        {
            if (App.EmployeeDict.TryGetValue(personnelCode, out var employee))
            {
                ConsoleService.WriteLine(employee.SignIn());
            }
        }
        else
        {
            ConsoleService.WriteLine("Invalid Maintenance Code.");
        }
        PersonnelNumberTextBox.Text = "";
    }

    private void signOutButton_Click(object sender, RoutedEventArgs e)
    {
        if (int.TryParse(PersonnelNumberTextBox.Text, out var personnelCode))
        {
            if (App.EmployeeDict.TryGetValue(personnelCode, out var employee))
            {
                ConsoleService.WriteLine( employee.SignOut() );
            }
        }
        else
        {
            ConsoleService.WriteLine("Invalid MaintenanceCode.");
        }
        PersonnelNumberTextBox.Text = "";
    }

    private async void debugOneButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            //var success = await LogImportExportService.GetRosterDate();
            //Console.WriteLine(success);
            LogImportExportService.GetRosterDate();
            //_ = ScheduleService.Task1();
            //Console.WriteLine();

            //var nightshifttable = LogImportExportService.ExportShiftLogWithSignInStatus(App.EmployeeDict, "NS");
            //Console.WriteLine(nightshifttable);

            //await ExcelService.NightShiftCrossoverAsync();
            //Console.ForegroundColor = ConsoleColor.Red;
            //Console.WriteLine("\nDebug");

            //Console.WriteLine(AppState.CurrentDate);
            //Console.WriteLine(AppState.Today);

            //Console.WriteLine(AppState.DayShiftStart);
            //Console.WriteLine(AppState.DayShiftEnd);

            //Console.WriteLine(AppState.NightShiftStart);
            //Console.WriteLine(AppState.NightShiftEnd);

            //Console.WriteLine();
            //Console.ResetColor();

            //Console.WriteLine(@"MaintenanceCodes Start");
            //foreach (var code in App.MaintenanceCodes)
            //{
            //    Console.WriteLine(code);
            //}
            //Console.WriteLine(@"Maintenance Codes End");
            //Console.WriteLine();
        }
        catch (Exception ex)
        {
            throw; // TODO handle exception
        }
    }
    private async void debugTwoButton_Click(object sender, RoutedEventArgs e)
    {
        Console.ForegroundColor = ConsoleColor.Blue;
        Console.WriteLine();
        //Console.WriteLine(@"Export Log Button");
        
        await LogImportExportService.SaveEmployeeDictionaryAsync(AppState.RosterStateFullPath);

        var files = FileHandling.GetFilesInFolder(AppState.StateFolder);
        var mostRecentRoster = FileHandling.GetMostRecentValidRosterFile(files);
        foreach (var file in files)
        {
            Console.WriteLine(file);
        }
        Console.WriteLine(mostRecentRoster);
        //LogImportExportService.ExportAdaptiveCardFromTemplateAsync(App.EmployeeDict, "DS");
        //Console.WriteLine(dayShiftLog);
        Console.ResetColor();
        Console.WriteLine();
    }

    // ================================ TASK BUTTONS ================================
   private async void taskOneButton_Click(object sender, RoutedEventArgs e)
    {
        Console.WriteLine(@"========== 5:00 AM ==========");
        try
        {
            _ = ScheduleService.Task1();
        }
        catch (Exception ex)
        {
            throw; // TODO handle exception
        }
    }

    private void taskTwoButton_Click(object sender, RoutedEventArgs e)
    {
        Console.WriteLine(@"========== 6:10 AM ==========");
        try
        {
            _ = ScheduleService.Task2();
        }
        catch (Exception ex)
        {
            throw; // TODO handle exception
        }
    }

    private void taskThreeButton_Click(object sender, RoutedEventArgs e)
    {
        Console.WriteLine(@"========== 7:00 AM ==========");
        try
        {
            _ = ScheduleService.Task3();
        }
        catch (Exception ex)
        {
            throw; // TODO handle exception
        }
    }
    
    private void taskFourButton_Click(object sender, RoutedEventArgs e)
    {
        Console.WriteLine(@"========== 6:10 PM ==========");
        try
        {
            _ = ScheduleService.Task4();
        }
        catch (Exception ex)
        {
            throw; // TODO handle exception
        }
    }

    private void taskFiveButton_Click(object sender, RoutedEventArgs e)
    {
        Console.WriteLine(@"========== 7:PM AM ==========");
        try
        {
            _ = ScheduleService.Task5();
        }
        catch (Exception ex)
        {
            throw; // TODO handle exception
        }
    }

    private void taskSixButton_Click(object sender, RoutedEventArgs e)
    {
        Console.WriteLine(@"========== 10:30 AM ==========");
        try
        {
            _ = ScheduleService.Task6();
        }
        catch (Exception ex)
        {
            throw; // TODO handle exception
        }
    }
    // ================================ TASK BUTTONS ================================

    private async Task ShowMessage(string title, string message)
    {
        var dialog = new ContentDialog
        {
            Title = title,
            Content = message,
            CloseButtonText = "OK",
            XamlRoot = XamlRoot // Set the XamlRoot to the current page's XamlRoot
        };

        // Show the dialog and wait for AppState.StartUpFunctionsComplete to be true
        var dialogTask = dialog.ShowAsync();
        while (!AppState.StartUpFunctionsComplete) { await Task.Delay(100); } // Check every 100ms

        // Close the dialog once AppState.StartUpFunctionsComplete is true
        //Log.Information("Closing Content Dialog");
        dialog.Hide();
        await dialogTask;
        //PersonnelNumberTextBox.Focus(FocusState.Programmatic);
    }

    private void PersonnelNumberTextBox_KeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (e.Key == VirtualKey.Enter)
        {
            if (int.TryParse(PersonnelNumberTextBox.Text, out var personnelCode))
            {
                // Got a valid int.
                if (App.EmployeeDict.TryGetValue(personnelCode, out var employee))
                {
                    var now = DateTime.Now;
                    string message;

                    // Define a cooldown period of 10 seconds for individual employees
                    var cooldownPeriod = TimeSpan.FromSeconds(10);

                    // Check if the last action for this specific employee was within the cooldown period
                    var isWithinCooldownPeriod = (employee.SignInTime.HasValue && (now - employee.SignInTime.Value) < cooldownPeriod)
                                                  || (employee.SignOutTime.HasValue && (now - employee.SignOutTime.Value) < cooldownPeriod);

                    if (isWithinCooldownPeriod)
                    {
                        message = $"{employee.Name} attempted action too soon. Please wait a few seconds before trying again.";
                    }
                    else
                    {
                        // Determine if it’s a valid sign-in time for the employee's shift
                        var isValidDayShiftSignIn = employee.ShiftType == "DS" && now.Hour >= 4 && now.Hour < 16;
                        var isValidNightShiftSignIn = employee.ShiftType == "NS" && (now.Hour >= 16 || now.Hour < 4);

                        if (employee.SignInTime.HasValue && !employee.SignOutTime.HasValue)
                        {
                            // Already signed in and it's not a valid sign-in time, so sign out
                            message = employee.SignOut();
                        }
                        else if (isValidDayShiftSignIn || isValidNightShiftSignIn)
                        {
                            // Valid sign-in time for shift, so sign in
                            message = employee.SignIn();
                        }
                        else
                        {
                            // Invalid sign-in time
                            message = "Invalid sign-in time for shift. Please try again during the appropriate hours.";
                        }
                    }

                    ConsoleService.WriteLine(message);
                }
                else
                {
                    ConsoleService.WriteLine("Invalid Maintenance Number.");
                }
            }
            else
            {
                // Didn't get a valid int.
                ConsoleService.WriteLine("Invalid Maintenance Number.");
            }

            // Clear the input
            PersonnelNumberTextBox.Text = "";

            // Optionally, prevent the default behavior of the Enter key
            e.Handled = true;
        }
    }
}