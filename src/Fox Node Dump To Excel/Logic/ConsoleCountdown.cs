using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;

namespace Fox_Node_Dump_Parser.Logic
{
    public class ConsoleCountdown
    {
        private readonly System.Threading.Timer timer;
        private int left;
        private int top;
        private readonly TimeSpan duration;
        private readonly Stopwatch stopwatch;
        public bool IsRunning => stopwatch.IsRunning;
        public bool CancellationRequested { get; set; }

        public ConsoleCountdown(TimeSpan duration)
        {
            this.duration = duration; 
            this.timer = new System.Threading.Timer(
                callback: new System.Threading.TimerCallback(TimerDisplayCountdown),
                state: null,
                dueTime: Timeout.Infinite,
                period: Timeout.Infinite);
            stopwatch = new Stopwatch();
        }

        public void Start()
        {
            Console.WriteLine("\nRight-click on tab icon get 'Export Text' option ... ");
            stopwatch.Restart();
            string preface = "Countdown to application closing: ";
            Console.Write(preface);
            var cursor = Console.GetCursorPosition();
            left = cursor.Left;
            top = cursor.Top;
            Console.WriteLine();
            timer.Change(dueTime: 0, period: 1000);

            while (!CancellationRequested && stopwatch.Elapsed < duration)
            {
                Task.Delay(10).Wait();
            }

            stopwatch.Stop();
            timer.Change(dueTime: Timeout.Infinite, period: Timeout.Infinite);
            timer.Dispose();
            Console.WriteLine();

        }

        public void Stop()
        {
            lock (this)
            {
                CancellationRequested = true;
            }
        }

        private void TimerDisplayCountdown(object? timerState)
        {
            Console.SetCursorPosition(left, top);   
            Console.Write($"{(duration - stopwatch.Elapsed):mm\\:ss}");
        }

       
    }
}
