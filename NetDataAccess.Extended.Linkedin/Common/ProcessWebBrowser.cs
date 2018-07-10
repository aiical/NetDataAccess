using NetDataAccess.Base.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace NetDataAccess.Extended.Linkedin.Common
{
    public class ProcessWebBrowser
    {
        [DllImport("shell32.dll")]
        static extern IntPtr ShellExecute(IntPtr hwnd, string lpOperation, string lpFile, string lpParameters, string lpDirectory, WebBrowserShowCommands nShowCmd);

        //清除IE所有访问痕迹
        public static void ClearWebBrowserTracks()
        {
            ShellExecute(IntPtr.Zero, "open", "rundll32.exe", " InetCpl.cpl,ClearMyTracksByProcess 255", "", WebBrowserShowCommands.SHOW_NO_GUI);
        }
        //清除IE Cookie
        public static void ClearWebBrowserCookie()
        {
            ShellExecute(IntPtr.Zero, "open", "rundll32.exe", " InetCpl.cpl,ClearMyTracksByProcess 2", "", WebBrowserShowCommands.SHOW_NO_GUI);
        }

        public static void AutoScroll(IRunWebPage runPage, WebBrowser webBrowser, int toPos, int maxStepLength, int minStepSleep, int maxStepSleep)
        {
            int pos = 0;
            Random random = new Random(DateTime.Now.Millisecond);
            while (pos < toPos)
            {
                int randomValue = random.Next(maxStepLength);
                pos += randomValue;
                runPage.InvokeScrollDocumentMethod(webBrowser, new Point(pos, pos));
                ProcessThread.SleepRandom(minStepSleep, maxStepSleep);
            }
        }
    }
}
