using Microsoft.UI.Dispatching;
using ScannerGUIv3.Core;
using Timer = System.Timers.Timer;
using ScannerGUIv3.Definitions;

namespace ScannerGUIv3.Services
{
    public class ScheduleService
    {
        private readonly List<string> personnelCodes = App.PersonnelCodes;
        public static Dictionary<int, Employee> EmployeeDict = App.EmployeeDict;

        // Constructor to initialize the scheduleTimes array
        private readonly TimeSpan[] scheduleTimes;

        // Timer to schedule tasks
        private readonly Timer timer;
        public ScheduleService()
        {
            //var now = DateTime.Now;
            scheduleTimes = new TimeSpan[]
            {
                new(4, 10, 0),
                new(6, 10, 0),
                new(7, 0, 0),
                new(18, 10, 0),
                new(19, 0, 0),
            };

            // This one is for debugging
            // Initialize the scheduleTimes array to execute 10 times, each 30 seconds apart
            scheduleTimes = new TimeSpan[10];
            var startTime = DateTime.Now.TimeOfDay;
            for (var i = 0; i < scheduleTimes.Length; i++)
            {
                scheduleTimes[i] = startTime.Add(TimeSpan.FromSeconds(10 * i));
            }

            // Initialize the timer
            timer = new Timer();

            // Schedule the first task
            ScheduleNextTask();
        }

        // Method to calculate the time until the next scheduled task
        private TimeSpan GetNextScheduledTime(DateTime now)
        {
            foreach (var time in scheduleTimes)
            {
                var next = now.Date + time;
                if (next > now)
                {
                    return next - now;
                }
            }
            // If all times are in the past, schedule the next task for tomorrow at the first time
            return (now.Date.AddDays(1) + scheduleTimes[0]) - now;
        }

        // Method to perform the scheduled operation
        private void PerformScheduledOperation()
        {
            var now = DateTime.Now.TimeOfDay;
            foreach (var time in scheduleTimes)
            {
                if (now >= time && now < time.Add(TimeSpan.FromMinutes(1)))
                {
                    if (time == new TimeSpan(4, 10, 0))
                    {
                        // Download roster files & initialize employee dictionary
                        // Set various variables
                        Task.Run(() => Task1());
                    }
                    else if (time == new TimeSpan(6, 10, 0))
                    {
                        Task.Run(() => Task2());
                    }
                    else if (time == new TimeSpan(7, 0, 0))
                    {
                        Task.Run(() => Task3());
                    }
                    else if (time == new TimeSpan(18, 10, 0))
                    {
                        Task.Run(() => Task4());
                    }
                    else if (time == new TimeSpan(19, 0, 0))
                    {
                        Task.Run(() => Task5());
                    }
                    break;
                }
            }
        }

        // Method to schedule the next task based on the current time
        private void ScheduleNextTask()
        {
            var now = DateTime.Now;
            var timeUntilNextTask = GetNextScheduledTime(now);

            // Initialize the timer with the calculated interval
            timer.Interval = timeUntilNextTask.TotalMilliseconds;
            timer.Elapsed += (sender, e) =>
            {
                // Stop the timer temporarily
                timer.Stop();

                // Check if DispatcherQueue is available
                var dispatcherQueue = DispatcherQueue.GetForCurrentThread();
                if (dispatcherQueue != null)
                {
                    // Enqueue the scheduled operation to run on the DispatcherQueue
                    dispatcherQueue.TryEnqueue(() => PerformScheduledOperation());
                }
                else
                {
                    // Perform the operation directly if no DispatcherQueue is available
                    PerformScheduledOperation();
                }

                // Reschedule the timer for the next time
                ScheduleNextTask();
            };
            timer.Start();
        }
        // Example tasks
        private async Task Task1() // 4:10 AM
        {
            Console.WriteLine(@"Task 1 executed at " + DateTime.Now);

            // Download the roster
            await Task.Run(() => LogImportExportService.DownloadExcelFileAsync(AppState.ResourcesExcelFolderPath, "Roster"));
            // Initialize the employee codes from HTTP
            await Task.Run(() => ExcelService.InitializeEmployeeCodesAsyncHTTP(personnelCodes, AppState.ResourcesMasterExcelPath));
            // Populate the dictionary
            await Task.Run(() => ExcelService.PopulateEmployeeDictionaryUsingXML(EmployeeDict, AppState.ResourcesOnSiteExcelPath));
            // Trim the dictionary
            await Task.Run(() => ExcelService.TrimEmployeeDictionaryAsync(EmployeeDict, App.PersonnelCodes));


            Console.WriteLine(@"Task 1 completed at " + DateTime.Now);
        }

        private static async Task Task2() // 6:10 AM
        {
            Console.WriteLine(@"Task 2 executed at " + DateTime.Now);
            // Generate the shiftLog for the concluding night shift
            var shiftLog = LogImportExportService.ExportShiftLog(EmployeeDict, "NS");
            // Export the shiftLog to Teams
            var success = await LogImportExportService.ExportToWeb(LogImportExportService.teamsUrl, shiftLog);
        }

        private static async Task Task3() // 7:00 am
        {
            Console.WriteLine(@"Task 3 executed at " + DateTime.Now);
            // Add your task 3 logic here
        }

        private static async Task Task4() // 6:10 PM
        {
            Console.WriteLine(@"Task 4 executed at " + DateTime.Now);
            // Generate the shiftLog for the concluding night shift
            var shiftLog = LogImportExportService.ExportShiftLog(EmployeeDict, "DS");
            // Export the shiftLog to Teams
            var success = await LogImportExportService.ExportToWeb(LogImportExportService.teamsUrl, shiftLog);
            if (success)
            {
            }
        }

        private static async Task Task5() // 7:00 PM
        {
            Console.WriteLine(@"Task 5 executed at " + DateTime.Now);
            // Add your task 5 logic here
        }
    }
}