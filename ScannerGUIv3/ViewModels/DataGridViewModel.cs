using System.Collections.ObjectModel;

using CommunityToolkit.Mvvm.ComponentModel;

using ScannerGUIv3.Contracts.ViewModels;
using ScannerGUIv3.Core.Contracts.Services;
using ScannerGUIv3.Core.Models;

namespace ScannerGUIv3.ViewModels;

public partial class DataGridViewModel : ObservableRecipient, INavigationAware
{
    private readonly ISampleDataService _sampleDataService;

    public ObservableCollection<SampleOrder> Source { get; } = new ObservableCollection<SampleOrder>();

    public DataGridViewModel(ISampleDataService sampleDataService)
    {
        _sampleDataService = sampleDataService;
    }

    public async void OnNavigatedTo(object parameter)
    {
        Source.Clear();

        // TODO: Replace with real data.
        Console.WriteLine("Break to inspect data.");
        var data = await _sampleDataService.GetGridDataAsync();

        foreach (var item in data)
        {
            Source.Add(item);
        }
        Console.WriteLine("Investigate source and data.");
    }

    public void OnNavigatedFrom()
    {
    }
}
