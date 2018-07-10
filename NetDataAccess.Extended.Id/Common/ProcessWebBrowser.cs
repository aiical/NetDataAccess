using NetDataAccess.Base.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace NetDataAccess.Extended.Taobao.Common
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

        public static void AutoScroll(IRunWebPage runPage, WebBrowser webBrowser, int toPosX, int toPosY, int maxStepLength, int minStepSleep, int maxStepSleep)
        {
            int posX = 0;
            int posY = 0;
            Random random = new Random(DateTime.Now.Millisecond);
            while (posX < toPosX || posY< toPosY)
            {
                int randomValue = random.Next(maxStepLength);
                posX = toPosX <= posX ? toPosX : posX + randomValue;
                posY = toPosY <= posY ? toPosY : posY + randomValue;
                runPage.InvokeScrollDocumentMethod(webBrowser, new Point(posX, posY));
                ProcessThread.SleepRandom(minStepSleep, maxStepSleep);
            }
        }
    }
}
