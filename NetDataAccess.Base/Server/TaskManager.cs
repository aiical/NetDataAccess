using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.Server;
using NetDataAccess.Base.Task;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace NetDataAccess.Base.Server
{
    /// <summary>
    /// 任务管理器
    /// </summary>
    public class TaskManager
    {
        #region MainRunningContainer
        private static IMainRunningContainer _RunningContainer = null;
        public static IMainRunningContainer RunningContainer
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
                    StartTaskMonitor();
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

        private static Dictionary<string, string> _CurrentTaskSteps = new Dictionary<string,string>();
        public static Dictionary<string, string> CurrentTaskSteps
        {
            get
            {
                return _CurrentTaskSteps;
            }
        }

        /// <summary>
        /// 可以新启动的Task数量
        /// </summary>
        /// <returns></returns>
        private static int GetIdleRunningCount()
        {
            int idleCount = SysConfig.MaxAliveTaskNum - CurrentTaskSteps.Count;
            return idleCount > 0 ? idleCount : 0;
        }

        public static void RemoveCurrentTask(string stepId)
        {
            lock (currentTaskLocker)
            {
                CurrentTaskSteps.Remove(stepId);
            }
        }
        public static void AddCurrentTask(string stepId, string taskId)
        {
            lock (currentTaskLocker)
            {
                CurrentTaskSteps.Add(stepId, taskId);
            }
        }
        #endregion

        #region 监视功能，自动启动爬取任务
        private static void StartTaskMonitor()
        {
            Thread monitorThread = new Thread(new ThreadStart(RunTaskMonitor));
            monitorThread.Start();
        } 
        #endregion

        #region 用于监视的线程执行的函数
        private static void RunTaskMonitor()
        {
            //将之前Running但是未完成的Step，置为Waiting状态
            ResetRunningSteps();


            while (1 == 1)
            {
                StartNewTaskSteps();
                Thread.Sleep(SysConfig.TaskMonitorInterval);
            }
        }
        #endregion

        #region 检测可以新启动的任务数，并启动任务
        private static void ResetRunningSteps()
        {
            TaskDataProcessor.ResetRunningSteps();
        }
        #endregion

        #region 检测可以新启动的任务数，并启动任务
        private static void StartNewTaskSteps()
        {
            int newTaskStepCount = GetIdleRunningCount();
            if (newTaskStepCount > 0)
            {
                List<Task_Step> steps = TaskDataProcessor.GetWaitingStepsInRunningTasks(newTaskStepCount);
                newTaskStepCount = newTaskStepCount - steps.Count;
                if (newTaskStepCount > 0)
                {
                    List<Task_Step> stepsInWaitingTasks = TaskDataProcessor.GetWaitingStepsInWaitingTasks(newTaskStepCount);
                    steps.AddRange(stepsInWaitingTasks);
                }

                foreach (Task_Step step in steps)
                {
                    try
                    {
                        TaskManager.AddCurrentTask(step.Id, step.TaskId);

                        TaskDataProcessor.UpdateDataAfterBeginStep(step.Id, step.TaskId);

                        TaskManager.RunningContainer.InvokeRunTask(step.GroupName, step.ProjectName, step.ListFilePath, step.InputDir, step.MiddleDir, step.OutputDir, step.Parameters, step.Id, true, false);
                    }
                    catch (Exception ex)
                    {
                        TaskDataProcessor.UpdateDataAfterEndStep(step.Id, step.TaskId, false, ex.Message);
                        continue;
                    }
                }
            }
        }
        #endregion
        
        #region 输出爬取是否成功到文件
        public static void ExportGrabResultFlag(string stepId, string projectName, bool succeed, string msg)
        {
            if (stepId != null && stepId.Length != 0)
            {
                string taskId = CurrentTaskSteps[stepId];
                TaskDataProcessor.UpdateDataAfterEndStep(stepId, taskId, succeed, msg);  
                TaskManager.RemoveCurrentTask(stepId);
            }
        }
        #endregion

        #region 启动任务
        public static void StartTask(string groupName, string projectName, string listFilePath, string parameters, string taskId, bool popPrompt)
        {
            StartTask(groupName, projectName, listFilePath, "", "", "", parameters, taskId, popPrompt);
        }
        #endregion

        #region 启动任务
        public static void StartTask(string groupName, string projectName, string listFilePath, string inputDir, string middleDir, string outputDir, string parameters, string taskId, bool popPrompt)
        {
            //启动后续任务
            try
            {
                bool autoRun = CommonUtil.IsNullOrBlank(taskId) || taskId == "_" ? false : true;
                taskId = CommonUtil.IsNullOrBlank(taskId) ? "_" : taskId;
                TaskManager.RunningContainer.InvokeRunTask(groupName, projectName, listFilePath, inputDir, middleDir, outputDir, parameters, taskId, autoRun, popPrompt);
            }
            catch (Exception ex)
            {
                throw new Exception("启动任务失败.ProjectName = " + projectName, ex);
            }
        }
        #endregion
    }
}
