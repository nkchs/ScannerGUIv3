namespace ScannerGUIv3.Core;

public class AppState
{
    private static readonly AppState _instance = new();
    public static AppState Instance => _instance;

    // ########## Variable Declarations ########## //
    // Time variables //
    public static DateTime currentDate = DateTime.Now;
    //public static Calendar calendar = CultureInfo.CurrentCulture.Calendar;

    public static DateTime today = DateTime.Today;
    // Define start and end times for day and night shifts
    public static DateTime dayShiftStart = today.AddHours(6);   // 6am on the same day
    public static DateTime dayShiftEnd = today.AddHours(18);    // 6pm on the same day

    public static DateTime nightShiftStart = today.AddHours(18); // 6pm on the same day
    public static DateTime nightShiftEnd = today.AddDays(1).AddHours(6); // 6am on the following day

    public static bool PersonnelCodesLoaded
    {
        get => _personnelCodesLoaded;
        set
        {
            _personnelCodesLoaded = value;
            Console.WriteLine($"PersonnelCodesLoaded set to: {value}");
        }
    }
    private static bool _personnelCodesLoaded;

    public static bool EmployeeDictionaryLoaded
    {
        get => _employeeDictionaryLoaded;
        set
        {
            _employeeDictionaryLoaded = value;
            Console.WriteLine($"EmployeeDictionaryLoaded set to: {value}");
        }
    }
    private static bool _employeeDictionaryLoaded;


    // Private constructor to prevent instantiation
    private AppState()
    {
    }
}