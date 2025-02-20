using System;
using System.Timers;
using Microsoft.UI.Dispatching;
using Timer = System.Timers.Timer;

namespace ScannerGUIv3.Services
{
    public class ScheduleService
    {
        private Timer timer;
        private readonly TimeSpan[] scheduleTimes =
        {
            new(15, 55, 0),
            new(15, 56, 0),
            new(15, 57, 0),
            new(15, 58, 0),
            new(15, 59, 0),
        };

        public void SetupDailyScheduler()
        {
            ScheduleNextTask();
        }

        private void ScheduleNextTask()
        {
            var now = DateTime.Now;
            var timeUntilNextTask = GetNextScheduledTime(now);
            timer = new Timer(timeUntilNextTask.TotalMilliseconds);
            timer.Elapsed += (sender, e) =>
            {
                timer.Stop();
                var dispatcherQueue = DispatcherQueue.GetForCurrentThread();
                if (dispatcherQueue != null)
                {
                    dispatcherQueue.TryEnqueue(() => PerformScheduledOperation());
                }
                else
                {
                    PerformScheduledOperation();
                }
                ScheduleNextTask();
            };
            timer.Start();
        }

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
            return (now.Date.AddDays(1) + scheduleTimes[0]) - now;
        }

        private static void PerformScheduledOperation()
        {
            Console.WriteLine("Scheduled operation executed at " + DateTime.Now);
        }
    }
}
