using Microsoft.Office.Interop.Excel;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using ScannerGUIv3.ViewModels;
using ScannerGUIv3.Helpers;
using System.Runtime.InteropServices;
using Application = Microsoft.UI.Xaml.Application;
using ScannerGUIv3.Definitions;
using Microsoft.UI.Xaml.Input;
using Windows.System;
using ScannerGUIv3.Services;
using Microsoft.UI.Xaml.Controls;
using ScannerGUIv3.Core;

namespace ScannerGUIv3.Views;

public sealed partial class MainPage : Microsoft.UI.Xaml.Controls.Page
{
    List<string> personnelCodes = ((App)Application.Current).personnelCodes;

    public MainViewModel ViewModel
    {
        get;
    }


    public MainPage() // 
    {
        ViewModel = App.GetService<MainViewModel>();
        //Console.WriteLine("Initializing Main Page.");
        InitializeComponent();
        ConsoleService.Initialize(ConsoleOutput); // Initialize with the console TextBox
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        PersonnelNumberTextBox.Focus(FocusState.Programmatic);
    }

    //static void ExcelLoader()
    //{
    //    Console.WriteLine("Excel Loader Called");
    //    //var excel_ = new Microsoft.Office.Interop.Excel.Application
    //    //{
    //    //    //Visible = false,
    //    //    Visible = true,
    //    //};

    //    //var resourcesOneSiteExcelUrl = @"https://newcrestmining-my.sharepoint.com/personal/nic_chase_newcrest_com_au/Documents/Documents/Projects/Scanner/ResourceOnSite_20240620043001.xlsx";
    //    ////string _excelURLOne = @"https://newcrestmining.sharepoint.com/:x:/r/teams/TelferEngineeringReliabilityGovernance/Shared%20Documents/General/Projects/Nic%20Chase/ResourceOnSite.xlsx";
    //    ////string _excelURLTwo = @"https://newcrestmining.sharepoint.com/:x:/r/teams/TelferMaint-Mill/Shared%20Documents/Attendance%20Register/FPM%20Daily%20Sign%20On/2024/Week%2043.xlsm?d=wc07871ef3cd04c2293ada7bcff29cc6d&csf=1&web=1&e=DORZmS";
    //    ////string _excelURLTwo = @"https://newcrestmining.sharepoint.com/:x:/r/teams/TelferMaint-Mill/Shared%20Documents/Attendance%20Register/FPM%20Daily%20Sign%20On/2024/Week%2043.xlsm";
    //    ////_excelURLOne = @"C:\Users\ChaseN\ResourceOnSite_20240620043001.xlsx";
    //    ////_excelURLOne = @"C:\Users\nicch\source\repos\nkchs\ScannerGUIv3\ResourceOnSite.xlsx";
    //    //var _excelURLTwo = resourcesOneSiteExcelUrl;
    //    //var _excelURLOne = resourcesOneSiteExcelUrl;

    //    ////Workbook excelWorkbook = excel_.Workbooks.Open(_excelURLOne, ReadOnly: true);
    //    //var excelWorkbook = excel_.Workbooks.Open(_excelURLTwo, ReadOnly: true);
    //    //var excelWorksheet = (Worksheet)excelWorkbook.Sheets[3];
    //    //var excelRange = excelWorksheet.UsedRange;
    //    //Console.WriteLine("Excel address: " + _excelURLOne);

    //    //var maxRow = excelRange.Rows.Count;
    //    //var maxCol = excelRange.Columns.Count;
    //    //Console.WriteLine("Rows: " + maxRow);
    //    //Console.WriteLine("Columns: " + maxCol);

    //    //excelWorkbook.Close(false, null, null);
    //    //excel_.Quit();
    //    //Marshal.ReleaseComObject(excelWorkbook);
    //    //Marshal.ReleaseComObject(excel_);
    //}

    private void signInButton_Click(object sender, RoutedEventArgs e)
    {
        if (int.TryParse(PersonnelNumberTextBox.Text, out var personnelCode))
        {
            if (App.EmployeeDict.TryGetValue(personnelCode, out var employee))
            {
                ConsoleService.WriteLine( employee.SignIn() );
            }
        }
        else
        {
            ConsoleService.WriteLine("Invalid Personnel Code.");
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
            ConsoleService.WriteLine("Invalid Personnel Code.");
        }
        PersonnelNumberTextBox.Text = "";
    }

    private void debugButton_Click(object sender, RoutedEventArgs e)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("\nDebug");

        Console.WriteLine(AppState.currentDate);
        Console.WriteLine(AppState.today);

        Console.WriteLine(AppState.dayShiftStart);
        Console.WriteLine(AppState.dayShiftEnd);

        Console.WriteLine(AppState.nightShiftStart);
        Console.WriteLine(AppState.nightShiftEnd);

        Console.WriteLine();

        Console.ResetColor();

        Console.WriteLine("Personnel Codes Start");
        foreach (var code in personnelCodes)
        {
            Console.WriteLine(code);
        }
        Console.WriteLine("Personnel Codes End");
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
                    ConsoleService.WriteLine("Invalid Personnel Number.");
                }
            }
            else
            {
                // Didn't get a valid int.
                ConsoleService.WriteLine("Invalid Personnel Number.");
            }

            // Clear the input
            PersonnelNumberTextBox.Text = "";

            // Optionally, prevent the default behavior of the Enter key
            e.Handled = true;
        }
    }
}