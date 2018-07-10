using NetDataAccess.Base.DLL;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.Server;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using log4net;

namespace NetDataAccess.Base.Log
{
    /// <summary>
    /// 执行任务日志
    /// </summary>
    public class RunTaskLog
    {
        #region 构造函数
        public RunTaskLog(string taskId)
        { 
            string logFilePath = Path.Combine(TaskManager.TaskFileDir, "log_" + taskId + ".txt"); 
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
        private object locker = new object();
        public void AddLog(string msg, LogLevelType logLevel, bool immediatelyShow)
        {
            lock (locker)
            {
                LogWriter.WriteLine(msg);
                if (immediatelyShow)
                {
                    LogWriter.Flush();
                }
            }
        }
        public void AddLog(string msg)
        {
            lock (locker)
            {
                LogWriter.WriteLine(msg); 
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
