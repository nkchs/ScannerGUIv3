using Microsoft.Office.Interop.Excel;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using ScannerGUIv3.ViewModels;
using ScannerGUIv3.Helpers;
using System.Runtime.InteropServices;
using Application = Microsoft.UI.Xaml.Application;

namespace ScannerGUIv3.Views;

public sealed partial class MainPage : Microsoft.UI.Xaml.Controls.Page
{
    public MainViewModel ViewModel
    {
        get;
    }


    public MainPage() // 
    {
        ViewModel = App.GetService<MainViewModel>();

        Console.WriteLine("Initializing Main Page.");

        InitializeComponent();
    }


    static void ExcelLoader()
    {
        var excel_ = new Microsoft.Office.Interop.Excel.Application
        {
            //Visible = false,
            Visible = true,
        };

        var resourcesOneSiteExcelUrl = @"https://newcrestmining-my.sharepoint.com/personal/nic_chase_newcrest_com_au/Documents/Documents/Projects/Scanner/ResourceOnSite_20240620043001.xlsx";
        //string _excelURLOne = @"https://newcrestmining.sharepoint.com/:x:/r/teams/TelferEngineeringReliabilityGovernance/Shared%20Documents/General/Projects/Nic%20Chase/ResourceOnSite.xlsx";
        //string _excelURLTwo = @"https://newcrestmining.sharepoint.com/:x:/r/teams/TelferMaint-Mill/Shared%20Documents/Attendance%20Register/FPM%20Daily%20Sign%20On/2024/Week%2043.xlsm?d=wc07871ef3cd04c2293ada7bcff29cc6d&csf=1&web=1&e=DORZmS";
        //string _excelURLTwo = @"https://newcrestmining.sharepoint.com/:x:/r/teams/TelferMaint-Mill/Shared%20Documents/Attendance%20Register/FPM%20Daily%20Sign%20On/2024/Week%2043.xlsm";
        //_excelURLOne = @"C:\Users\ChaseN\ResourceOnSite_20240620043001.xlsx";
        //_excelURLOne = @"C:\Users\nicch\source\repos\nkchs\ScannerGUIv3\ResourceOnSite.xlsx";
        var _excelURLTwo = resourcesOneSiteExcelUrl;
        var _excelURLOne = resourcesOneSiteExcelUrl;

        //Workbook excelWorkbook = excel_.Workbooks.Open(_excelURLOne, ReadOnly: true);
        var excelWorkbook = excel_.Workbooks.Open(_excelURLTwo, ReadOnly: true);
        var excelWorksheet = (Worksheet)excelWorkbook.Sheets[3];
        var excelRange = excelWorksheet.UsedRange;
        Console.WriteLine("Excel address: " + _excelURLOne);


        var maxRow = excelRange.Rows.Count;
        var maxCol = excelRange.Columns.Count;
        Console.WriteLine("Rows: " + maxRow);
        Console.WriteLine("Columns: " + maxCol);


        excelWorkbook.Close(false, null, null);
        excel_.Quit();
        Marshal.ReleaseComObject(excelWorkbook);
        Marshal.ReleaseComObject(excel_);
    }


    private void signInButton_Click(object sender, RoutedEventArgs e)
    {

        signInButton.Content = "Clicked";
        ExcelLoader();
        var PersonnelCode = PersonnelNumberTextBox.Text;
        Console.WriteLine("Personnel Code: " + PersonnelCode);
        signInButton.Content = "Sign In";
        PersonnelNumberTextBox.Text = "";
    }


    private void debugButton_Click(object sender, RoutedEventArgs e)
    {
        //var diccccc = App.employeeDict;
        Console.WriteLine("Debug");

        Console.WriteLine(App.excelWeeklyAddress);
        Console.WriteLine(App.excelResourceAddress);

        Console.WriteLine("Recorded DateTime: " + App.currentDate);
        Console.WriteLine("Current DateTime: " + DateTime.Now);


        Console.WriteLine(App.currentDate.Day);

        Console.WriteLine("Week Number: " + App.weekNumber);
        App.weekNumber = 44;

        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("Week Number: " + App.weekNumber);
        Console.ResetColor();


        Console.WriteLine("");

        // Old Code
        //Console.WriteLine("Bounds:" + App.MainWindow.Bounds.ToString());
    }
}