using Microsoft.UI.Xaml.Controls;

using ScannerGUIv3.ViewModels;

namespace ScannerGUIv3.Views;

public sealed partial class BlankPage : Page
{
    public BlankViewModel ViewModel
    {
        get;
    }

    public BlankPage()
    {
        ViewModel = App.GetService<BlankViewModel>();
        InitializeComponent();
    }
}