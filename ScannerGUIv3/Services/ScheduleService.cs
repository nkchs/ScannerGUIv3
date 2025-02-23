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
            // This one is live
            // Initialize the scheduleTimes array with the specified times
            scheduleTimes = new TimeSpan[]
            {
                new(4, 10, 0),
                new(6, 10, 0),
                new(7, 0, 0),
                new(18, 10, 0),
                new(19, 0, 0)
            };


            // This one is for debugging
            // Initialize the scheduleTimes array to execute 10 times, each 30 seconds apart
            scheduleTimes = new TimeSpan[10];
            var startTime = DateTime.Now.TimeOfDay;
            for (var i = 0; i < scheduleTimes.Length; i++)
            {
                scheduleTimes[i] = startTime.Add(TimeSpan.FromSeconds(10 * i));
            }


            timer = new Timer(); // Initialize the timer to avoid CS8618 error
            // Schedule the first task
            ScheduleNextTask();
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
                        Task1();
                    }
                    else if (time == new TimeSpan(6, 10, 0))
                    {
                        Task2();
                    }
                    else if (time == new TimeSpan(7, 0, 0))
                    {
                        Task3();
                    }
                    else if (time == new TimeSpan(18, 10, 0))
                    {
                        Task4();
                    }
                    else if (time == new TimeSpan(19, 0, 0))
                    {
                        Task5();
                    }
                    break;
                }
            }
        }

        // Example tasks
        private void Task1()
        {
            Console.WriteLine("Task 1 executed at " + DateTime.Now);
            // Add your task 1 logic here
        }

        private void Task2()
        {
            Console.WriteLine("Task 2 executed at " + DateTime.Now);
            // Add your task 2 logic here
        }

        private void Task3()
        {
            Console.WriteLine("Task 3 executed at " + DateTime.Now);
            // Add your task 3 logic here
        }

        private void Task4()
        {
            Console.WriteLine("Task 4 executed at " + DateTime.Now);
            // Add your task 4 logic here
        }

        private void Task5()
        {
            Console.WriteLine("Task 5 executed at " + DateTime.Now);
            // Add your task 5 logic here
        }
    }
}