using System;
using System.Collections.Generic;
using System.Windows.Forms;
using NetDataAccess.Edit;
using NetDataAccess.Main;
using System.IO;
using log4net.Repository;
using NetDataAccess.Base.DLL;
using NetDataAccess.Base.Server;

namespace NetDataAccess
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            //网络连接设置，允许最近网络连接数为512
            System.Net.ServicePointManager.DefaultConnectionLimit = 512;

            //加载日志配置
            ILoggerRepository rep = log4net.LogManager.CreateRepository("grab");
            string log4netFilePath = Path.Combine(Path.GetDirectoryName(Application.StartupPath), "Files/Config/log4net.config");
            log4net.Config.XmlConfigurator.ConfigureAndWatch(rep, new System.IO.FileInfo(log4netFilePath));

            FormMain runForm = new FormMain();

            Application.Run(runForm); 
        } 
    }
}
