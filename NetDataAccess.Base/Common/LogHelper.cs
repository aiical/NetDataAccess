using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms; 

namespace NetDataAccess.Base.Common
{
    /// <summary>
    /// 日志记录类
    /// </summary>
    public class LogHelper
    {
        #region 程序相关的临时文件夹位置
        /// <summary>
        /// 程序相关的临时文件夹位置
        /// </summary>
        private static string TempDirectoryPath = Path.Combine(Path.GetTempPath(), "NcpNetDataAccess");
        #endregion

        #region 日志存储的位置
        /// <summary>
        /// 日志存储的位置
        /// </summary>
        private static string LogDirectoryPath = Path.Combine(TempDirectoryPath, "Log");
        #endregion

        #region 锁
        /// <summary>
        /// 锁
        /// </summary>
        private static object locker = new object();
        #endregion

        #region 日志写入流
        /// <summary>
        /// 日志写入流
        /// </summary>
        private static StreamWriter SW = null;
        #endregion

        #region 写入日志
        /// <summary>
        /// 写入日志
        /// </summary>
        /// <param name="errMessage"></param>
        public static void WriteMessage(string errMessage)
        {
            try
            {
                string filePath = null;
                DateTime time = DateTime.Now;
                lock (locker)
                {
                    if (!Directory.Exists(TempDirectoryPath))
                    {
                        Directory.CreateDirectory(TempDirectoryPath);
                        if (!Directory.Exists(LogDirectoryPath))
                        {
                            Directory.CreateDirectory(LogDirectoryPath);
                        }
                    }

                    filePath = Path.Combine(TempDirectoryPath, time.ToString("yyyy-MM-dd") + ".log");
                    if (!File.Exists(filePath))
                    {
                        StreamWriter newSW = new StreamWriter(filePath, true);
                        if (SW != null)
                        {
                            SW.Close();
                            SW.Dispose();
                        }
                        SW = newSW;
                    }
                    else
                    {
                        if (SW == null)
                        {
                           SW = new StreamWriter(filePath, true);
                        }
                    }
                }
                 
                SW.Write(time.ToString("yyyy-MM-dd HH:mm:ss:ffff - "));
                SW.WriteLine(errMessage);  
            }
            catch (Exception ex)
            {
                MessageBox.Show("日志记录出错." + ex.Message);
            }
        }
        #endregion
    }
}
