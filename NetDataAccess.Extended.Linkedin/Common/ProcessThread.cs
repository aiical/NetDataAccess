using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace NetDataAccess.Extended.Linkedin.Common
{
    public class ProcessThread
    {
        public static void SleepRandom(int min, int max)
        {
            Random r = new Random(DateTime.Now.Millisecond);
            int sleepTime = min + r.Next(max - min);
            Thread.Sleep(sleepTime);
        }
    }
}
