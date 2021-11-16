using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TricentisExtensions.Modules.WindowForms
{
    public class CancelDialog : Window
    {
        public CancelDialog() : base("Cancel Dialog")
        {
        }

        public CancelDialog(int seconds) : base("Cancel Dialog", "", seconds)
        {
        }

        public bool WaitWindowDisappear(int timeout)
        {
            var stopwatch = new Stopwatch();
            
            bool exists = true;
            stopwatch.Start();


            while (exists)
            {
                var check = WinExists(1);

                if (check == 0)
                {
                    exists = false;
                }

                double seconds = stopwatch.Elapsed.TotalSeconds;

                if (seconds > timeout)
                {
                    stopwatch.Stop();
                    throw new TimeoutException($"Operation is cancelled by timeout. Waited {seconds} seconds for condition.");
                }
            }

            return exists;
        }
    }
}
