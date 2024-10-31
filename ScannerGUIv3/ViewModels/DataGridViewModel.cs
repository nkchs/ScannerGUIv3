using System.Collections.ObjectModel;

using CommunityToolkit.Mvvm.ComponentModel;

using ScannerGUIv3.Contracts.ViewModels;
using ScannerGUIv3.Core.Contracts.Services;
using ScannerGUIv3.Core.Models;
using ScannerGUIv3.Definitions;

namespace ScannerGUIv3.ViewModels;


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