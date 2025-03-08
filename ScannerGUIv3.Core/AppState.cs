using Microsoft.Identity.Client.Extensibility;

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


    // File Variables

    //public static string ResourcesExcelFolderPath
    //{
    //    get; set;
    //}
    //public static string ResourcesOnSiteExcelPath
    //{
    //    get; set;
    //}
    //public static string ResourcesMasterExcelPath
    //{
    //    get; set;
    //}

    private static string _resourcesExcelFolderPath;
    public static string ResourcesExcelFolderPath
    {
        get => _resourcesExcelFolderPath;
        set
        {
            _resourcesExcelFolderPath = value;
            ResourcesOnSiteExcelPath = _resourcesExcelFolderPath + @"/Roster.xlsx";
            ResourcesMasterExcelPath = _resourcesExcelFolderPath + @"/SRF195 Profile Master.xlsx";
        }
    }

    public static string ResourcesOnSiteExcelPath
    {
        get; private set;
    }
    public static string ResourcesMasterExcelPath
    {
        get; private set;
    }


    public static bool PersonnelCodesLoaded
    {
        get => _personnelCodesLoaded;
        set
        {
            _personnelCodesLoaded = value;
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"PersonnelCodesLoaded set to: {value}");
            Console.ResetColor();
        }
    }

    private static bool _personnelCodesLoaded;

    public static bool EmployeeDictionaryLoaded
    {
        get => _employeeDictionaryLoaded;
        set
        {
            _employeeDictionaryLoaded = value;
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"EmployeeDictionaryLoaded set to: {value}");
            Console.ResetColor();
        }
    }

    private static bool _employeeDictionaryLoaded;

    public static bool EmployeeDictionaryTrimmed
    {
        get => _employeeDictionaryTrimmed;
        set
        {
            _employeeDictionaryTrimmed = value;
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"EmployeeDictionaryTrimmed set to: {value}");
            Console.ResetColor();
        }
    }

    private static bool _employeeDictionaryTrimmed;

    public static bool EmployeeDictionaryRefreshed
    {
        get => _employeeDictionaryRefreshed;
        set
        {
            _employeeDictionaryRefreshed = value;
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"EmployeeDictionaryRefreshed set to: {value}");
            Console.ResetColor();
        }
    }

    private static bool _employeeDictionaryRefreshed;

    // Private constructor to prevent instantiation
    private AppState()
    {
    }
}