using System.Collections.ObjectModel;

using CommunityToolkit.Mvvm.ComponentModel;

using ScannerGUIv3.Contracts.ViewModels;
using ScannerGUIv3.Core.Contracts.Services;
using ScannerGUIv3.Core.Models;
using ScannerGUIv3.Definitions;

namespace ScannerGUIv3.ViewModels;

//public partial class DataGridViewModel : ObservableRecipient, INavigationAware
//{
//    private readonly ISampleDataService _sampleDataService;

//    public ObservableCollection<SampleOrder> Source { get; } = new ObservableCollection<SampleOrder>();

//    public DataGridViewModel(ISampleDataService sampleDataService)
//    {
//        _sampleDataService = sampleDataService;
//    }

//    public async void OnNavigatedTo(object parameter)
//    {
//        Source.Clear(); // What is source?

//        // TODO: Replace with real data.
//        Console.WriteLine("Break to inspect data.");
//        var data = await _sampleDataService.GetGridDataAsync(); // Type appears to be IEnumerable List

//        foreach (var item in data)
//        {
//            Source.Add(item);
//        }
//        Console.WriteLine("Investigate source and data.");
//    }

//    public void OnNavigatedFrom()
//    {
//    }
//}


public partial class DataGridViewModel : ObservableRecipient, INavigationAware
{
    public ObservableCollection<Employee> Employees { get; } = new ObservableCollection<Employee>();

    public DataGridViewModel()
    {
        // Initialize with existing employee data from App.EmployeeDict
        LoadEmployeesFromDictionary();
    }

    public void OnNavigatedTo(object parameter)
    {
        // Clear the current list to prevent duplication
        Employees.Clear();

        // Populate Employees with items from App.EmployeeDict
        LoadEmployeesFromDictionary();
    }

    public void OnNavigatedFrom()
    {
    }

    private void LoadEmployeesFromDictionary()
    {
        foreach (var employee in App.EmployeeDict.Values)
        {
            Employees.Add(employee);
        }
    }
}