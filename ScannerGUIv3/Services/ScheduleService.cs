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
        private bool _isExecuting = false;

        private static TimeSpan[]? _scheduleTimes;

        // Timer to schedule tasks
        private readonly Timer _timer;
        public ScheduleService()
        {
            _scheduleTimes =
            [
                new TimeSpan(4, 10, 0),  // Task1
                new TimeSpan(6, 10, 0),  // Task2
                new TimeSpan(7, 0, 0),   // Task3
                new TimeSpan(10, 30, 0), // Task6
                new TimeSpan(18, 10, 0), // Task4
                new TimeSpan(19, 0, 0)   // Task5
            ];
            
            // Initialize the timer
            _timer = new Timer();

            // Schedule the first task
            ScheduleNextTask();
        }

        // Method to perform the scheduled operation
        private void PerformScheduledOperation()
        {
            if (_isExecuting)
            {
                return; // Prevent concurrent executions
            }

            _isExecuting = true;
            try
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
            finally
            {
                _isExecuting = false;
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


        // ==================== TASKS ============================================================================================================
        public static async Task Task1()
        {
            Log.Information("Task 1: Starting roster processing at {Time}", DateTime.Now);

            try
            {
                //Log.Verbose("============= TASK 1 STARTED =============");

                Log.Debug("Preparing previous night shift data");
                await Task.Run(ExcelService.GeneratePreviousNightShiftAsync);
                Log.Debug("Previous night shift data prepared");

                const int maxRetryAttempts = 3;
                const int retryDelayMilliseconds = 5000;
                var attempt = 0;
                var downloadSuccess = false;

                Log.Debug("Attempting to download roster file");
                while (attempt < maxRetryAttempts && !downloadSuccess)
                {
                    attempt++;
                    downloadSuccess = await LogImportExportService.DownloadExcelFileAsync(
                        AppState.ResourcesExcelFolderPath, "Roster", LogImportExportService.WorkforceReportDownloadUri);

                    if (downloadSuccess)
                    {
                        Log.Information("Successfully downloaded roster file on attempt {Attempt}", attempt);
                    }
                    else
                    {
                        Log.Warning("Failed to download roster file on attempt {Attempt}. Retrying in {Delay} seconds",
                            attempt, retryDelayMilliseconds / 1000);
                        await Task.Delay(retryDelayMilliseconds);
                    }
                }

                if (downloadSuccess)
                {
                    Log.Debug("Loading employee data into dictionary");
                    await Task.Run(() => ExcelService.PopulateEmployeeDictionaryUsingXml(AppState.ResourcesOnSiteExcelPath));
                    Log.Debug("Employee data loaded into dictionary");

                    Log.Debug("Preparing next night shift data");
                    await Task.Run(ExcelService.GenerateNextNightShiftAsync);
                    Log.Debug("Next night shift data prepared");

                    Log.Debug("Inserting current night shift data");
                    await Task.Run(ExcelService.InsertCurrentNightShiftAsync);
                    Log.Debug("Current night shift data inserted");

                    Log.Information("Task 1: Completed successfully at {Time}", DateTime.Now);
                }
                else
                {
                    Log.Warning("Task 1: Roster download failed after {MaxAttempts} attempts", maxRetryAttempts);

                    const int maxRecursiveAttempts = 5;
                    const int delayMilliseconds = 300000;

                    if (AppState.TaskOneAttempt < maxRecursiveAttempts)
                    {
                        AppState.TaskOneAttempt++;
                        Log.Information("Initiating retry {Attempt} of {MaxAttempts} after {Delay} minutes",
                            AppState.TaskOneAttempt, maxRecursiveAttempts, delayMilliseconds / 60000);
                        await Task.Delay(delayMilliseconds);
                        Log.Debug("Retrying Task 1");
                        await Task1();
                    }
                    else
                    {
                        Log.Error("Task 1: Exhausted maximum retry attempts ({MaxAttempts})", maxRecursiveAttempts);
                        AppState.TaskOneAttempt = 0;
                    }
                }

                //Log.Verbose("============= TASK 1 FINISHED =============");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Task 1: Critical error during roster processing");
                throw;
            }
        }

        //public static async Task Task1()
        //{
        //    Console.WriteLine();
        //    //Log.Verbose(@"============= TASK 1 EXECUTED =============");
        //    //Log.Information("Task 1: Starting night shift export process at {Time}", DateTime.Now);

        //    // Remove the CURRENT night shift employees from the dictionary
        //    await Task.Run(ExcelService.GeneratePreviousNightShiftAsync);

        //    const int maxRetryAttempts = 3; // Maximum number of retry attempts
        //    const int retryDelayMilliseconds = 5000; // Delay between retries (in milliseconds)
        //    var attempt = 0;
        //    var downloadSuccess = false;

        //    // Retry mechanism
        //    while (attempt < maxRetryAttempts && !downloadSuccess)
        //    {
        //        attempt++;
        //        downloadSuccess =
        //            await LogImportExportService.DownloadExcelFileAsync(AppState.ResourcesExcelFolderPath,
        //                "Roster", LogImportExportService.WorkforceReportDownloadUri);

        //        if (downloadSuccess)
        //        {
        //            Log.Information($"Download succeeded on attempt {attempt}.");
        //        }
        //        else
        //        {
        //            Log.Warning(
        //                $"Download failed on attempt {attempt}. Retrying in {retryDelayMilliseconds / 1000} seconds...");
        //            await Task.Delay(retryDelayMilliseconds);
        //        }
        //    }
        //    if (downloadSuccess)    // Continue with the remaining tasks if the download was successful
        //    {
        //        // Populate the employee dictionary
        //        await Task.Run(() =>
        //            ExcelService.PopulateEmployeeDictionaryUsingXml(AppState.ResourcesOnSiteExcelPath));

        //        // Remove the UPCOMING night shift employees from the dictionary
        //        await Task.Run(ExcelService.GenerateNextNightShiftAsync);
        //        // Insert the PREVIOUS night shift employees into the dictionary
        //        await Task.Run(ExcelService.InsertCurrentNightShiftAsync);

        //        Log.Information("Maintenance Codes & Dictionary & Trim & Next Night Shift");
        //    }
        //    else
        //    {
        //        // Handle the case where all retry attempts fail
        //        Log.Error("Failed to download the roster after maximum retry attempts.");
        //        // Add fallback or recovery logic here (e.g., notifying the user, alternative actions, etc.)
        //        const int maxRecursiveAttempts = 5;
        //        const int delayMilliseconds = 300000; // 5 minutes

        //        if (AppState.TaskOneAttempt < maxRecursiveAttempts)
        //        {
        //            AppState.TaskOneAttempt++; // Increment the attempt counter
        //            Log.Information($"Waiting {delayMilliseconds / 60000} minutes before retry attempt {AppState.TaskOneAttempt} of {maxRecursiveAttempts}");
        //            await Task.Delay(delayMilliseconds); // Wait 5 minutes
        //            await Task1(); // Recursively call Task1
        //        }
        //        else
        //        {
        //            Log.Error($"Maximum retry attempts ({maxRecursiveAttempts}) reached");
        //            AppState.TaskOneAttempt = 0; // Reset the counter after max attempts
        //        }
        //    }
        //    Log.Verbose(@"============= TASK 1 COMPLETED ============");
        //    Console.WriteLine();
        //}


        public static async Task Task2() // 6:10 AM Export the concluding night shift
        {
            Log.Information("Task 2: Starting night shift export process at {Time}", DateTime.Now);

            try
            {
                Log.Debug("Sending night shift employee data to {Email}", "nic.chase@greatland.com.au");
                var success = await LogImportExportService.SendEmployeeDataAsync("nic.chase@greatland.com.au", "NS");

                if (success)
                {
                    Log.Information("Night shift data exported successfully");

                    Log.Debug("Removing previous night shift records");
                    await Task.Run(ExcelService.RemovePreviousNightShiftAsync);
                    Log.Debug("Previous night shift records removed");

                    Log.Debug("Inserting next night shift records");
                    await Task.Run(ExcelService.InsertNextNightShiftAsync);
                    Log.Debug("Next night shift records inserted");

                    Log.Information("Task 2: Completed successfully at {Time} - Night shift transition complete", DateTime.Now);
                }
                else
                {
                    Log.Warning("Task 2: Failed to export night shift data - Skipping shift transition");
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Task 2: Unexpected error during night shift export process");
                throw;
            }
        }


        public static async Task Task3() // 7:00 am Export the starting dayshift
        {
            Log.Verbose(@"Task 3 Executed");
            //var success = await LogImportExportService.ExportAdaptiveCardFromTemplateAsync(App.EmployeeDict, "DS");
            var success = await LogImportExportService.SendEmployeeDataAsync("nic.chase@greatland.com.au", "DS"); // This is the email export option.
            Log.Verbose(success ? @"Task 3 Completed" : @"Task 3 Failed");
        }

        public static async Task Task6() // 10:30 AM Export the current dayshift to capture flights in
        {
            Log.Verbose(@"Task 6 Executed");
            //var success = await LogImportExportService.ExportAdaptiveCardFromTemplateAsync(App.EmployeeDict, "DS");
            var success = await LogImportExportService.SendEmployeeDataAsync("nic.chase@greatland.com.au", "DS"); // This is the email export option.
            Log.Verbose(success ? @"Task 6 Completed" : @"Task 6 Failed");
        }

        public static async Task Task4() // 6:10 PM
        {
            Log.Verbose(@"Task 4 Executed");
            try
            {
                //var success = await LogImportExportService.ExportAdaptiveCardFromTemplateAsync(App.EmployeeDict, "NS");
                var success = await LogImportExportService.SendEmployeeDataAsync("nic.chase@greatland.com.au", "NS"); // This is the email export option.
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