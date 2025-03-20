using System.Net.NetworkInformation;

namespace ScannerGUIv3.Core;

public class AppState
{
    public static bool Logging;

    //private static readonly AppState Instance = new();

    // ########## Variable Declarations ########## //
    // Time variables //
    public static DateTime CurrentDate = DateTime.Now;
    public static DateTime Today = DateTime.Today;

    // Define start and end times for day and night shifts
    public static DateTime DayShiftStart = Today.AddHours(6);   // 6am on the same day
    public static DateTime DayShiftEnd = Today.AddHours(18);    // 6pm on the same day

    public static DateTime NightShiftStart = Today.AddHours(18); // 6pm on the same day
    public static DateTime NightShiftEnd = Today.AddDays(1).AddHours(6); // 6am on the following day


    // File Variables
    private static string _resourcesExcelFolderPath;
    public static string ResourcesExcelFolderPath
    {
        get => _resourcesExcelFolderPath;
        set
        {
            _resourcesExcelFolderPath = value;
            ResourcesOnSiteExcelPath = _resourcesExcelFolderPath + @"/Roster.xlsx";
            ResourcesMasterExcelPath = _resourcesExcelFolderPath + @"/SRF195 Profile Master.xlsx";
            LogFolder = _resourcesExcelFolderPath; // + @"/Logs";
            StateFolder = _resourcesExcelFolderPath; // + @"/State"; 
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

    public static string LogFolder
    {
        get;
        private set;
    }
    public static string StateFolder
    {
        get; private set;
    }

    // Bool flags
    public static bool StartUpFunctionsComplete = false;
    public static bool MaintenanceCodesLoaded
    {
        get => _maintenanceCodesLoaded;
        set
        {
            _maintenanceCodesLoaded = value;
            if (!Logging)
            {
                return;
            }
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"MaintenanceCodesLoaded set to: {value}");
            Console.ResetColor();
        }
    }
    private static bool _maintenanceCodesLoaded;

    public static bool EmployeeDictionaryLoaded
    {
        get => _employeeDictionaryLoaded;
        set
        {
            _employeeDictionaryLoaded = value;
            if (!Logging)
            {
                return;
            }
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
            if (!Logging)
            {
                return;
            }
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
            if (!Logging)
            {
                return;
            }
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"EmployeeDictionaryRefreshed set to: {value}");
            Console.ResetColor();
        }
    }
    private static bool _employeeDictionaryRefreshed;

    public static bool TrimShiftRequired = true; // Whether or not trim is required

    // Task Variables
    public static int TaskOneAttempt = 0;
    public static int TaskTwoAttempt = 0;
    public static int TaskThreeAttempt = 0;
    public static int TaskFourAttempt = 0;
    public static int TaskFiveAttempt = 0;

    // Private constructor to prevent instantiation
    private AppState()
    {
    }
}

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