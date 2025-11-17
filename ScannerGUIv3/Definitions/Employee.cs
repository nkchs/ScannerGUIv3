namespace ScannerGUIv3.Definitions;
public class Employee
{
    public int EmployeeNumber { get; set; }

    public string? Name { get; set; }

    public string? ShiftType
    {
        get => _shiftType;// Try to get today's shift from the ShiftSchedule dictionary//if (ShiftSchedule.TryGetValue(DateTime.Today, out var shift))//{//    return NormalizeShiftType(shift);//}//return null;

        set => _shiftType = NormalizeShiftType(value);
    }

    private string? _shiftType;

    // Helper function to normalize shift values
    private static string? NormalizeShiftType(string? value)
    {
        //Console.WriteLine(value);
        if (string.IsNullOrEmpty(value))
        {
            return "OS";
        }
        else if (value.StartsWith("D", StringComparison.OrdinalIgnoreCase))
        {
            return "DS";  // Day Shift
        }
        else if (value.StartsWith("N", StringComparison.OrdinalIgnoreCase))
        {
            return "NS";  // Night Shift
        }
        else if (value.Equals("RR", StringComparison.OrdinalIgnoreCase))
        {
            return "OS";  // Rostered Offsite
        }
        return "OS";  // Invalid shift type, set to "OS Offsite"
    }

    public Dictionary<DateTime, string> ShiftSchedule { get; set; } = [];

    public DateTime? SignInTime { get; private set; }

    public DateTime? SignOutTime { get; private set; }

    // Default constructor
    public Employee()
    {
    }

    public Employee(int employeeNumber, string name)
    {
        EmployeeNumber = employeeNumber;
        Name = name;
    }

    //Method to sign in, setting the SignInTime to now
    public string SignIn()
    {
        if (SignInTime.HasValue)
        {
            return "Employee has already signed in. Sign out before signing in again.";
        }

        SignInTime = DateTime.Now;
        return EmployeeNumber + " " + Name + " Signed In @ " + FormattedSignInTime;
    }

    public string SignOut()
    {
        if (SignInTime == null)
        {
            return "Previous sign in record not found. Logged to file.";
        }
        else
        {
            SignOutTime = DateTime.Now;
            return EmployeeNumber + " " + Name + " Signed Out @ " + FormattedSignOutTime;
        }
    }
    
    public string ToAsciiTableRow()
    {
        const int signpadding = 11;
        var name = Name?.PadRight(28) ?? "No Name".PadRight(28);
        var id = EmployeeNumber.ToString().PadRight(8);
        var signIn = SignInTime.HasValue ? FormattedSignInTime.PadRight(signpadding) : "No Sign In".PadRight(signpadding);
        var signOut = SignOutTime.HasValue ? FormattedSignOutTime.PadRight(signpadding) : "No Sign Out".PadRight(signpadding);

        return $"| {name} | {id} | {signIn} | {signOut} |";
    }
    
    // Property to get formatted SignInTime
    public string FormattedSignInTime => SignInTime?.ToString("HH:mm") ?? "No Sign In";

    public string FormattedSignOutTime => SignOutTime?.ToString("HH:mm") ?? "No Sign Out";

    // RETIRED
    // RETIRED
    // Parameterized constructor
    //public Employee(int employeeNumber, string name, string shiftType, DateTime signInTime, DateTime signOutTime)
    //{
    //    EmployeeNumber = employeeNumber;
    //    Name = name;
    //    ShiftType = shiftType;
    //    SignInTime = signInTime;
    //    SignOutTime = signOutTime;
    //}

    //public Employee(int employeeNumber, string name, string shiftType)
    //{
    //    EmployeeNumber = employeeNumber;
    //    Name = name;
    //    ShiftType = shiftType;
    //    SignInTime = null;
    //    SignOutTime = null;
    //}

    //public string? ShiftType
    //{
    //    get => shiftType;
    //    set
    //    {
    //        if (string.IsNullOrEmpty(value))
    //        {
    //            shiftType = null;
    //        }
    //        else if (value.StartsWith("D", StringComparison.OrdinalIgnoreCase))
    //        {
    //            shiftType = "DS";  // Day Shift
    //        }
    //        else if (value.StartsWith("N", StringComparison.OrdinalIgnoreCase))
    //        {
    //            shiftType = "NS";  // Night Shift
    //        }
    //        else if (value.Equals("RR", StringComparison.OrdinalIgnoreCase))
    //        {
    //            shiftType = "RR";  // Rostered Offsite
    //        }
    //        else
    //        {
    //            shiftType = null;  // Invalid shift type, set to null
    //        }
    //    }
    //}

    //public string? ShiftType
    //{
    //    get
    //    {
    //        // Try to get today's shift from the ShiftSchedule dictionary
    //        if (ShiftSchedule.TryGetValue(DateTime.Today, out var shift))
    //        {
    //            return NormalizeShiftType(shift);
    //        }
    //        return null;
    //    }
    //    //set => shiftType = NormalizeShiftType(value);
    //    set
    //    {
    //        NormalizeShiftType(value);
    //    }
    //}
}
