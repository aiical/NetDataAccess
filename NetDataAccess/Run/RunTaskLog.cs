using NetDataAccess.Base.DLL;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Main;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NetDataAccess.Run
{
    /// <summary>
    /// 执行任务日志
    /// </summary>
    public class RunTaskLog
    {
        #region 构造函数
        public RunTaskLog(string taskId)
        {
            string logFilePath = Path.Combine(Path.Combine(TaskManager.TaskFileDir,taskId),"log.txt");
            TextWriter logWriter = File.CreateText(logFilePath);
            _LogWriter = logWriter;
        }
        #endregion

        #region LogWriter
        private TextWriter _LogWriter = null;
        private TextWriter LogWriter
        {
            get
            {
                return _LogWriter;
            }
        }
        #endregion


        #region 记录日志
        public void AddLog(string msg, LogLevelType logLevel, bool immediatelyShow)
        {
            LogWriter.WriteLine(msg);
            if (immediatelyShow)
            {
                LogWriter.Flush();
            }
        }
        #endregion

        #region 关闭
        public void Close()
        {
            if (LogWriter != null)
            {
                LogWriter.Flush();
                LogWriter.Close();
            }
        }
        #endregion
    }
}
