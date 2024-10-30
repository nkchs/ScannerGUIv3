using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Networking;

namespace ScannerGUIv3.Definitions;
public class Employee
{
    public int EmployeeNumber { get; set; }

    public string? Name { get; set; }

    public string? ShiftType { get; set; }

    public DateTime? SignInTime { get; set; }

    public DateTime? SignOutTime { get; set; }

    public DateTime? ArrivalDate { get; set; }

    public DateTime? DepartureDate { get; set; }

    // Default constructor
    public Employee()
    {
    }

    // Parameterized constructor
    public Employee(int employeeNumber, string name, string shiftType, DateTime signInTime, DateTime signOutTime)
    {
        EmployeeNumber = employeeNumber;
        Name = name;
        ShiftType = shiftType;
        SignInTime = signInTime;
        SignOutTime = signOutTime;
    }

    public Employee(int employeeNumber, string name, string shiftType)
    {
        EmployeeNumber = employeeNumber;
        Name = name;
        ShiftType = shiftType;
        SignInTime = null;
        SignOutTime = null;
    }
}
