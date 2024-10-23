using Microsoft.Office.Interop.Excel;
using Microsoft.UI.Xaml;
using ScannerGUIv3.ViewModels;

namespace ScannerGUIv3.Views;

public sealed partial class MainPage : Microsoft.UI.Xaml.Controls.Page
{
    public MainViewModel ViewModel
    {
        get;
    }

    public MainPage()
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

        string _excelURLOne = @"https://newcrestmining.sharepoint.com/:x:/r/teams/TelferEngineeringReliabilityGovernance/Shared%20Documents/General/Projects/Nic%20Chase/ResourceOnSite.xlsx";
        //_excelURLOne = @"C:\Users\ChaseN\ResourceOnSite_20240620043001.xlsx";
        //_excelURLOne = @"C:\Users\nicch\source\repos\nkchs\ScannerGUIv3\ResourceOnSite.xlsx";

        Workbook excelWorkbook = excel_.Workbooks.Open(_excelURLOne, ReadOnly: true);
        Worksheet excelWorksheet = (Worksheet)excelWorkbook.Sheets[3];
        Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;
        Console.WriteLine("Excel address: " + _excelURLOne);

        int maxRow = excelRange.Rows.Count;
        int maxCol = excelRange.Columns.Count;
        Console.WriteLine("Rows: " + maxRow);
        Console.WriteLine("Columns: " + maxCol);

        //excelWorkbook.Close(false, null, null);
        //excel_.Quit();
        //Marshal.ReleaseComObject(excelWorkbook);
        //Marshal.ReleaseComObject(excel_);
    }

    private void signInButton_Click(object sender, RoutedEventArgs e)
    {
        signInButton.Content = "Clicked";
        ExcelLoader();
        string PersonnelCode = PersonnelNumberTextBox.Text;
        Console.WriteLine("Personnel Code: " + PersonnelCode);
        signInButton.Content = "Sign In";
        PersonnelNumberTextBox.Text = "";
    }

    private void debugButton_Click(object sender, RoutedEventArgs e)
    {
        Console.WriteLine("Debug");
        Console.WriteLine("");
    }
}