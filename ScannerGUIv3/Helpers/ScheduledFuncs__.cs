using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.UI.Dispatching;
using Timer = System.Timers.Timer;

namespace ScannerGUIv3.Helpers;
public class ScheduledFuncs
{

    public required Timer timer;

    public void SetupDailyScheduler()
    {
        ScheduleNextTask();
    }


    public void ScheduleNextTask()
    {
        // 
        DateTime now = DateTime.Now;
        TimeSpan timeUntilNextTask = GetNextScheduledTime(now);
        // Set the timer to trigger at the calculated interval
        timer = new Timer(timeUntilNextTask.TotalMilliseconds);
        timer.Elapsed += (sender, e) =>
        {
            timer.Stop();  // Stop the timer temporarily

            // Check if DispatcherQueue is available
            var dispatcherQueue = DispatcherQueue.GetForCurrentThread();
            if (dispatcherQueue != null)
            {
                dispatcherQueue.TryEnqueue(() => PerformScheduledOperation());
            }
            else
            {
                // Perform operation directly if no DispatcherQueue is available
                PerformScheduledOperation();
            }

            // Reschedule the timer for the next time
            ScheduleNextTask();
        };
        timer.Start();
    }


    public TimeSpan GetNextScheduledTime(DateTime now)
    {
        foreach (var time in scheduleTimes)
        {
            //DateTime next = now.Date + now.TimeOfDay + time;
            DateTime next = now.Date + time;
            if (next > now) return next - now;
        }
        // If all times are in the past, the next scheduled time is tomorrow at the first time
        return (now.Date.AddDays(1) + scheduleTimes[0]) - now;
    }


    public void PerformScheduledOperation()
    {
        // Your task code here, which runs at 5:00 AM, 5:00 PM, and midnight
        Console.WriteLine("Scheduled operation executed at " + DateTime.Now);
    }
}
