//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using Windows.Networking;

namespace ScannerGUIv3.Definitions;
public class Employee
{
    public int EmployeeNumber { get; set; }

    public string? Name { get; set; }

    //public string? ShiftType { get; set; }
    private string? shiftType;

    public string? ShiftType
    {
        get => shiftType;
        set
        {
            if (string.IsNullOrEmpty(value))
            {
                shiftType = null;
            }
            else if (value.StartsWith("D", StringComparison.OrdinalIgnoreCase))
            {
                shiftType = "DS";  // Day Shift
            }
            else if (value.StartsWith("N", StringComparison.OrdinalIgnoreCase))
            {
                shiftType = "NS";  // Night Shift
            }
            else if (value.Equals("RR", StringComparison.OrdinalIgnoreCase))
            {
                shiftType = "RO";  // Rostered Offsite
            }
            else
            {
                shiftType = null;  // Invalid shift type, set to null
            }
        }
    }

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

    //Method to sign in, setting the SignInTime to now
    public string SignIn()
    {
        if (SignInTime.HasValue)
        {
            return "Already signed in. Sign out before signing in again.";
        }

        SignInTime = DateTime.Now;
        return EmployeeNumber + " " + Name + " Signed In @ " + FormattedSignInTime;
    }

    //public string SignIn()
    //{
    //    // Check if the employee is already signed in
    //    if (SignInTime.HasValue)
    //    {
    //        return "Already signed in. Sign out before signing in again.";
    //    }

    //    DateTime now = DateTime.Now;

    //    // Check valid sign-in times
    //    if (ShiftType == "DS" && (now.Hour < 4 || now.Hour >= 16))
    //    {
    //        return "Invalid sign-in time for Day Shift. Valid hours are 4 AM to 4 PM.";
    //    }
    //    else if (ShiftType == "NS" && (now.Hour >= 4 && now.Hour < 16))
    //    {
    //        return "Invalid sign-in time for Night Shift. Valid hours are 4 PM to 4 AM.";
    //    }

    //    SignInTime = now;
    //    return EmployeeNumber + " " + Name + " Signed In @ " + FormattedSignInTime;
    //}

    public string SignOut()
    {
        if (SignInTime == null)
        {
            return "Not signed in.";
        }
        else
        {
            SignOutTime = DateTime.Now;
            return EmployeeNumber + " " + Name + " Signed Out @ " + FormattedSignInTime;
        }
    }

    // Property to get formatted SignInTime
    public string FormattedSignInTime => SignInTime?.ToString("dd/MM/yyyy HH:mm") ?? "Not signed in.";

    // Property to get formatted SignOutTime
    public string FormattedSignOutTime => SignOutTime?.ToString("dd/MM/yyyy HH:mm") ?? "Not signed out.";
}
