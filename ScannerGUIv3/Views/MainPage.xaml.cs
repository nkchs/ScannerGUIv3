using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

using ScannerGUIv3.ViewModels;
using ScannerGUIv3.Helpers;
using ScannerGUIv3.Core.Helpers;
using Microsoft.Office.Interop.Excel;

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
        InitializeComponent();
    }

    static void ExcelLoader()
    {
        var excel = new Microsoft.Office.Interop.Excel.Application
        {
            Visible = false,
            // Visible = true,
        };

        Workbook excelWorkbook = excel.Workbooks.Open(@"C:\Users\ChaseN\ResourceOnSite_20240620043001.xlsx");

        Worksheet excelWorksheet = (Worksheet)excelWorkbook.Sheets[3];
        Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;

        int maxRow = excelRange.Rows.Count;
        int maxCol = excelRange.Columns.Count;
        Console.WriteLine("Rows: " + maxRow);
        Console.WriteLine("Columns: " + maxCol);
    }

    private void signInButton_Click(object sender, RoutedEventArgs e)
    {
        signInButton.Content = "Clicked";
        ExcelLoader();
        string PersonnelCode = PersonnelNumberTextBox.Text;

    }


}
