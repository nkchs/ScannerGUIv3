using Microsoft.UI.Xaml.Controls;

using ScannerGUIv3.ViewModels;

namespace ScannerGUIv3.Views;

public sealed partial class ContentGridPage : Page
{
    public ContentGridViewModel ViewModel
    {
        get;
    }

    public ContentGridPage()
    {
        ViewModel = App.GetService<ContentGridViewModel>();
        InitializeComponent();
    }
}