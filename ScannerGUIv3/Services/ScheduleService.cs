using System;
using System.Timers;
using Microsoft.UI.Dispatching;
using Timer = System.Timers.Timer;

namespace ScannerGUIv3.Services
{
    public class ScheduleService
    {
        // Timer to schedule tasks
        private Timer timer;

        // Constructor to initialize the scheduleTimes array
        private readonly TimeSpan[] scheduleTimes;

        public ScheduleService()
        {
            // Initialize the scheduleTimes array to execute 10 times, each 30 seconds apart
            scheduleTimes = new TimeSpan[10];
            var startTime = DateTime.Now.TimeOfDay;
            for (int i = 0; i < scheduleTimes.Length; i++)
            {
                scheduleTimes[i] = startTime.Add(TimeSpan.FromSeconds(10 * i));
            }
            ScheduleNextTask();
        }

        // Method to schedule the next task based on the current time
        private void ScheduleNextTask()
        {
            var now = DateTime.Now;
            var timeUntilNextTask = GetNextScheduledTime(now);

            // Initialize the timer with the calculated interval
            timer = new Timer(timeUntilNextTask.TotalMilliseconds);
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
            if (now >= new TimeSpan(3, 0, 0) && now < new TimeSpan(9, 0, 0))
            {
                morningUpdate();
            }
            else if (now >= new TimeSpan(18, 0, 0) && now < new TimeSpan(24, 0, 0))
            {
                eveningUpdate();
            }
            else
            {
                Console.WriteLine("No scheduled operation for this time.");
            }
        }

        // Example function to perform a morning update
        private void morningUpdate()
        {
            Console.WriteLine("Morning update executed at " + DateTime.Now);
            // Add your morning update logic here
        }

        // Example function to perform an evening update
        private void eveningUpdate()
        {
            Console.WriteLine("Evening update executed at " + DateTime.Now);
            // Add your evening update logic here
        }
    }
}
