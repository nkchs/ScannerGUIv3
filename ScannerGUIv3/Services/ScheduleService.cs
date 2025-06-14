using Microsoft.UI.Dispatching;
using ScannerGUIv3.Core;
using Timer = System.Timers.Timer;
using ScannerGUIv3.Definitions;
using Serilog;
using System.Net.Mail;
using DocumentFormat.OpenXml.Drawing;

namespace ScannerGUIv3.Services
{
    public class ScheduleService
    {
        // Constructor to initialize the scheduleTimes array
        //private readonly TimeSpan[] _scheduleTimes;
        private static TimeSpan[]? _scheduleTimes;

        // Timer to schedule tasks
        private readonly Timer _timer;
        public ScheduleService()
        {
            //var now = DateTime.Now;
            _scheduleTimes =
            [
                new TimeSpan(5, 00, 0),
                new TimeSpan(6, 30, 0),
                new TimeSpan(7, 0, 0),
                new TimeSpan(10, 30, 0), // New task at 10:30 AM
                new TimeSpan(18, 10, 0),
                new TimeSpan(19, 0, 0)
                //new(DateTime.Now.Hour, DateTime.Now.Minute+1, DateTime.Now.Second),
            ];


            // TURN ON FOR DEBUG
            // Initialize the scheduleTimes array to execute 10 times, each 30 seconds apart
            //scheduleTimes = new TimeSpan[10];
            //var startTime = DateTime.Now.TimeOfDay;
            //for (var i = 0; i < scheduleTimes.Length; i++)
            //{
            //    scheduleTimes[i] = startTime.Add(TimeSpan.FromSeconds(10 * i));
            //}
            // TURN ON FOR DEBUG


            // Initialize the timer
            _timer = new Timer();

            // Schedule the first task
            ScheduleNextTask();
        }

        // Method to perform the scheduled operation
        private void PerformScheduledOperation()
        {
            var now = DateTime.Now.TimeOfDay;
            foreach (var time in _scheduleTimes)
            {
                if (now >= time && now < time.Add(TimeSpan.FromMinutes(5)))
                {
                    if (time == new TimeSpan(4, 10, 0))
                    {
                        Task.Run(Task1);
                    }
                    else if (time == new TimeSpan(6, 10, 0))
                    {
                        Task.Run(Task2);
                    }
                    else if (time == new TimeSpan(7, 0, 0))
                    {
                        Task.Run(Task3);
                    }
                    else if (time == new TimeSpan(10, 30, 0)) // New task at 10:30 AM
                    {
                        Task.Run(Task6);
                    }
                    else if (time == new TimeSpan(18, 10, 0))
                    {
                        Task.Run(Task4);
                    }
                    else if (time == new TimeSpan(19, 0, 0))
                    {
                        Task.Run(Task5);
                    }
                    else
                    {
                        Log.Verbose("Alternative Task");
                    }
                    break;
                }
            }
        }

        // Method to calculate the time until the next scheduled task
        private TimeSpan GetNextScheduledTime(DateTime now)
        {
            foreach (var time in _scheduleTimes)
            {
                var next = now.Date + time;
                if (next > now)
                {
                    return next - now;
                }
            }
            // If all times are in the past, schedule the next task for tomorrow at the first time
            return (now.Date.AddDays(1) + _scheduleTimes[0]) - now;
        }


        // Method to schedule the next task based on the current time
        private void ScheduleNextTask()
        {
            var now = DateTime.Now;
            var timeUntilNextTask = GetNextScheduledTime(now);

            // Initialize the timer with the calculated interval
            _timer.Interval = timeUntilNextTask.TotalMilliseconds;
            _timer.Elapsed += (sender, e) =>
            {
                // Stop the timer temporarily
                _timer.Stop();

                // Check if DispatcherQueue is available
                var dispatcherQueue = DispatcherQueue.GetForCurrentThread();
                if (dispatcherQueue != null)
                {
                    // Enqueue the scheduled operation to run on the DispatcherQueue
                    dispatcherQueue.TryEnqueue(PerformScheduledOperation);
                }
                else
                {
                    // Perform the operation directly if no DispatcherQueue is available
                    PerformScheduledOperation();
                }

                // Reschedule the timer for the next time
                ScheduleNextTask();
            };
            _timer.Start();
        }

        // TASKS
        public static async Task Task1()
        {
            Console.WriteLine();
            Log.Verbose(@"============= TASK 1 EXECUTED =============");

            // Remove the CURRENT night shift employees from the dictionary
            await Task.Run(ExcelService.GeneratePreviousNightShiftAsync);

            // Check if the roster has been updated today
            if (await LogImportExportService.GetRosterDate()) // True if the roster has been updated today
            {
                const int maxRetryAttempts = 3; // Maximum number of retry attempts
                const int retryDelayMilliseconds = 5000; // Delay between retries (in milliseconds)
                var attempt = 0;
                var downloadSuccess = false;

                // Retry mechanism
                while (attempt < maxRetryAttempts && !downloadSuccess)
                {
                    attempt++;
                    downloadSuccess =
                        await LogImportExportService.DownloadExcelFileAsync(AppState.ResourcesExcelFolderPath,
                            "Roster");

                    if (downloadSuccess)
                    {
                        Log.Information($"Download succeeded on attempt {attempt}.");
                    }
                    else
                    {
                        Log.Warning(
                            $"Download failed on attempt {attempt}. Retrying in {retryDelayMilliseconds / 1000} seconds...");
                        await Task.Delay(retryDelayMilliseconds);
                    }
                }
                if (downloadSuccess)    // Continue with the remaining tasks if the download was successful
                {
                    // Populate the employee dictionary
                    await Task.Run(() =>
                        ExcelService.PopulateEmployeeDictionaryUsingXml(AppState.ResourcesOnSiteExcelPath));

                    // Remove the UPCOMING night shift employees from the dictionary
                    await Task.Run(ExcelService.GenerateNextNightShiftAsync);
                    // Insert the PREVIOUS night shift employees into the dictionary
                    await Task.Run(ExcelService.InsertCurrentNightShiftAsync);

                    Log.Information("Maintenance Codes & Dictionary & Trim & Next Night Shift");
                }
                else
                {
                    // Handle the case where all retry attempts fail
                    Log.Error("Failed to download the roster after maximum retry attempts.");
                    // Add fallback or recovery logic here (e.g., notifying the user, alternative actions, etc.)
                    const int maxRecursiveAttempts = 5;
                    const int delayMilliseconds = 300000; // 5 minutes

                    if (AppState.TaskOneAttempt < maxRecursiveAttempts)
                    {
                        AppState.TaskOneAttempt++; // Increment the attempt counter
                        Log.Information($"Waiting {delayMilliseconds / 60000} minutes before retry attempt {AppState.TaskOneAttempt} of {maxRecursiveAttempts}");
                        await Task.Delay(delayMilliseconds); // Wait 5 minutes
                        await Task1(); // Recursively call Task1
                    }
                    else
                    {
                        Log.Error($"Maximum retry attempts ({maxRecursiveAttempts}) reached");
                        AppState.TaskOneAttempt = 0; // Reset the counter after max attempts
                    }
                }
            }
            else
            {
                Log.Error("Roster Has Not Been Updated");
            }
            Log.Verbose(@"============= TASK 1 COMPLETED ============");
            Console.WriteLine();
        }

        public static async Task Task2() // 6:10 AM Export the concluding night shift
        {
            Log.Verbose(@"Task 2 Executed");
            var success = await LogImportExportService.ExportAdaptiveCardFromTemplateAsync(App.EmployeeDict, "NS");
            if (success)
            {
                // Remove the previous Night Shift
                await Task.Run(ExcelService.RemovePreviousNightShiftAsync);
                // Insert the next Night Shift
                await Task.Run(ExcelService.InsertNextNightShiftAsync);
            }
            Log.Verbose(success ? @"Task 2 Completed" : @"Task 2 Failed");
        }

        public static async Task Task3() // 7:00 am Export the starting dayshift
        {
            Log.Verbose(@"Task 3 Executed");
            var success = await LogImportExportService.ExportAdaptiveCardFromTemplateAsync(App.EmployeeDict, "DS");
            Log.Verbose(success ? @"Task 3 Completed" : @"Task 3 Failed");
        }

        public static async Task Task6() // 10:30 AM Export the current dayshift to capture flights in
        {
            Log.Verbose(@"Task 6 Executed");
            var success = await LogImportExportService.ExportAdaptiveCardFromTemplateAsync(App.EmployeeDict, "DS");
            Log.Verbose(success ? @"Task 6 Completed" : @"Task 6 Failed");
        }

        public static async Task Task4() // 6:10 PM
        {
            Log.Verbose(@"Task 4 Executed");
            try
            {
                var success = await LogImportExportService.ExportAdaptiveCardFromTemplateAsync(App.EmployeeDict, "NS");
                // Generate the crossover shiftLog
                if (success)
                {
                    await ExcelService.NightShiftCrossoverAsync();
                    App.EmployeeDict = App.EmployeeCrossoverDict;
                }
                AppState.TaskFourAttempt = 0; // Reset attempt counter on success
                Log.Verbose(@"Task 4 Completed");
            }
            catch (Exception ex)
            {
                Log.Warning($"Task 4 failed: {ex.Message}");
                const int maxRetryAttempts = 5;
                const int retryDelayMilliseconds = 300000; // 5 minutes

                if (AppState.TaskFourAttempt < maxRetryAttempts)
                {
                    AppState.TaskFourAttempt++; // Increment attempt counter
                    Log.Warning($"Task 4 failed. Attempt {AppState.TaskFourAttempt} of {maxRetryAttempts}. Retrying in {retryDelayMilliseconds / 60000} minutes...");
                    await Task.Delay(retryDelayMilliseconds);
                    await Task4(); // Retry the task
                }
                else
                {
                    Log.Error($"Task 4 failed after {maxRetryAttempts} attempts: {ex.Message}");
                    AppState.TaskFourAttempt = 0; // Reset counter after max attempts
                    throw; // Re-throw the exception after max attempts
                }
            }
        }

        public static async Task Task5() // 7:00 PM
        {
            Log.Verbose(@"Task 5 Executed");
        }
    }
}