using Microsoft.UI.Xaml.Controls;

using ScannerGUIv3.ViewModels;

namespace ScannerGUIv3.Views;

// TODO: Change the grid as appropriate for your app. Adjust the column definitions on DataGridPage.xaml.
// For more details, see the documentation at https://docs.microsoft.com/windows/communitytoolkit/controls/datagrid.
public sealed partial class DataGridPage : Page
{
    public DataGridViewModel ViewModel
    {
        get;
    }

    public DataGridPage()
    {
        ViewModel = App.GetService<DataGridViewModel>();
        //ViewModel = App.GetService<DataGridPersonnelModel>();
        Console.WriteLine("Initializing Data Grid Page.");
        InitializeComponent();
    }
}
