using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace NetDataAccess.Main
{
    /// <summary>
    /// 任务管理器
    /// </summary>
    public class TaskManager
    {
        #region MainRunningContainer
        private static MainRunningContainer _RunningContainer = null;
        public static MainRunningContainer RunningContainer
        {
            get
            {
                return _RunningContainer;
            }
            set
            {
                if (_RunningContainer == null)
                {
                    _RunningContainer = value;
                    StartMonitor();
                }
            }
        }
        #endregion

        #region 关闭任务UI
        public static void CloseTaskUI(string taskId)
        {
            RunningContainer.InvokeCloseTaskUI(taskId);
        }
        #endregion

        #region task启动监视线程轮询间隔时间（秒）
        private static int IntervalCheckNewTask = 3;
        #endregion

        #region Files文件夹
        private static string _FileDir = null;
        public static string FileDir
        {
            get
            {
                if (_FileDir == null)
                {
                    _FileDir = Path.Combine(Path.GetDirectoryName(Application.StartupPath), "Files");
                }
                return _FileDir;
            }
        }
        #endregion

        #region 输入文件路径
        private static string _InputFileDir = null;
        public static string InputFileDir
        {
            get
            {
                if (_InputFileDir == null)
                {
                    _InputFileDir = Path.Combine(FileDir, "Input");
                }
                return _InputFileDir;
            }
        }
        #endregion

        #region 输出文件路径
        private static string _OutputFileDir = null;
        public static string OutputFileDir
        {
            get
            {
                if (_OutputFileDir == null)
                {
                    _OutputFileDir = Path.Combine(FileDir, "Output");
                }
                return _OutputFileDir;
            }
        }
        #endregion

        #region 运行的任务文件路径
        private static string _TaskFileDir = null;
        public static string TaskFileDir  
        {
            get
            {
                if (_TaskFileDir == null)
                {
                    _TaskFileDir = Path.Combine(FileDir, "Task");
                }
                return _TaskFileDir;
            }
        } 
        #endregion

        #region 记录正在爬取的任务Id
        private static object currentTaskLocker = new object();

        private static string _CurrentTaskId = null;
        public static string CurrentTaskId
        {
            get
            {
                return _CurrentTaskId;
            }
        }

        /// <summary>
        /// 判断是否有任务正在执行
        /// </summary>
        /// <returns></returns>
        private static bool CheckNoneRunning()
        {
            return CurrentTaskId == null;
        }

        public static void CheckAndSetCurrentTask(string taskId)
        {
            lock (currentTaskLocker)
            {
                if (_CurrentTaskId == null || taskId == null)
                {
                    _CurrentTaskId = taskId; 
                }
                else
                {
                    throw new Exception("设置当前任务出错. NewTaskId = " + taskId + ", CurrentTaskId=" + _CurrentTaskId + ".");
                }
            }
        }
        #endregion

        #region 监视功能，自动启动爬取任务
        private static void StartMonitor()
        {
            Thread monitorThread = new Thread(new ThreadStart(MonitorThread));
            monitorThread.Start();
        } 
        #endregion

        #region 用于监视的线程执行的函数
        private static void MonitorThread()
        {
            while (1 == 1)
            {
                Thread.Sleep(IntervalCheckNewTask * 1000);
                MonitorInputFolder();
            }
        }
        #endregion

        #region 监视Input文件夹，如果发现新任务，那么将新任务文件夹放置到Task目录下,并启动任务
        private static void MonitorInputFolder()
        {
            if (CheckNoneRunning())
            {
                string inputDirs = Path.Combine(FileDir, "Input");
                string[] inputTaskDirs = Directory.GetDirectories(inputDirs);
                foreach (string inputTaskDir in inputTaskDirs)
                {
                    string inputErrorFilePath = Path.Combine(inputTaskDir, "error.xml");
                    //如果不存在错误提示文件，那么可以执行此任务
                    if (!File.Exists(inputErrorFilePath))
                    {
                        string inputFilePath = Path.Combine(inputTaskDir, "input.xml");
                        string taskId = Path.GetFileName(Path.GetDirectoryName(inputFilePath));
                        string groupName = "";
                        string projectName = "";
                        string parameter = "";
                        if (File.Exists(inputFilePath) || (!CommonUtil.IsFileInUse(inputFilePath)))
                        {
                            try
                            {
                                string taskDir = Path.Combine(TaskFileDir, taskId);

                                //将input文件夹里的内容放置到task文件夹里
                                if (Directory.Exists(taskDir))
                                {
                                    throw new Exception("已执行过相同的任务Id. TaskId=" + taskId);
                                }
                                Directory.Move(inputTaskDir, taskDir);

                                inputFilePath = Path.Combine(taskDir, "input.xml");
                                /*
                                 <Task 
                                    GroupName=""
                                    ProjectName=""
                                    Parameter=""
                                 ></Task>
                                 */
                                XmlDocument inputDoc = new XmlDocument();
                                inputDoc.Load(inputFilePath);
                                XmlElement rootNode = inputDoc.DocumentElement;
                                groupName = rootNode.Attributes["GroupName"].Value;
                                projectName = rootNode.Attributes["ProjectName"].Value;
                                parameter = rootNode.Attributes["Parameter"].Value;

                                TaskManager.CheckAndSetCurrentTask(taskId);

                                TaskManager.RunningContainer.InvokeRunTask(groupName, projectName, parameter, taskId);
                                break;
                            }
                            catch (Exception ex)
                            {
                                ExportGrabResultFlag(taskId, projectName, "001", CommonUtil.GetExceptionAllMessage(ex), DateTime.Now, DateTime.Now, true);
                                continue;
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
            }
        }
        #endregion


        #region 输出爬取是否成功到文件
        public static void ExportGrabResultFlag(string taskId, string projectName, string code, string msg, DateTime startTime, DateTime endTime, bool saveErrorToInputDir)
        {
            //将运行的结果状态信息保存到Output目录下
            string flagFilePath = Path.Combine(Path.Combine(OutputFileDir, taskId), "result.xml");
            CommonUtil.CreateFileDirectory(flagFilePath);

            XmlDocument flagDoc = new XmlDocument();
            flagDoc.LoadXml("<?xml version=\"1.0\" encoding=\"utf-8\" ?><Result></Result>");
            XmlNode rootNode = flagDoc.DocumentElement;

            XmlAttribute taskIdAttri = flagDoc.CreateAttribute("TaskId");
            taskIdAttri.Value = taskId;
            rootNode.Attributes.Append(taskIdAttri);

            XmlAttribute projectNameAttri = flagDoc.CreateAttribute("ProjectName");
            projectNameAttri.Value = projectName;
            rootNode.Attributes.Append(projectNameAttri);

            XmlAttribute codeAttri = flagDoc.CreateAttribute("Code");
            codeAttri.Value = code;
            rootNode.Attributes.Append(codeAttri);

            XmlAttribute msgAttri = flagDoc.CreateAttribute("Message");
            msgAttri.Value = msg;
            rootNode.Attributes.Append(msgAttri);

            XmlAttribute endTimeAttri = flagDoc.CreateAttribute("EndTime");
            endTimeAttri.Value = endTime.ToString("yyyy-MM-dd HH:mm:ss");
            rootNode.Attributes.Append(endTimeAttri);

            XmlAttribute startTimeAttri = flagDoc.CreateAttribute("StartTime");
            startTimeAttri.Value = endTime.ToString("yyyy-MM-dd HH:mm:ss");
            rootNode.Attributes.Append(startTimeAttri);

            flagDoc.Save(flagFilePath);

            if(saveErrorToInputDir)
            {
                string inputErrorFilePath = Path.Combine(Path.Combine(InputFileDir, taskId), "error.xml");
                File.Copy(flagFilePath, inputErrorFilePath);
            }

            //设置正在执行的任务为空
            TaskManager.CheckAndSetCurrentTask(null);
        }
        #endregion
    }
}
