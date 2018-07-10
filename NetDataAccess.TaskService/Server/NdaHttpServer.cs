using NetDataTaskManager.Manager;
using NetDataTaskManager.Task;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web;

namespace NetDataAccess.TaskService.Server
{ 
    public class NdaHttpServer : HttpServer
    {
        #region _DefaultServer
        private static NdaHttpServer _DefaultServer = null;        
        #endregion

        #region StartServer
        public static void StartServer(string ip, int port)
        {
            if (_DefaultServer == null)
            {
                _DefaultServer = new NdaHttpServer(ip, port);
            }
            if (!_DefaultServer.IsActive)
            {
                _DefaultServer.Start();
            }
        }
        #endregion

        #region StopServer
        public static void StopServer()
        {
            if (_DefaultServer != null && _DefaultServer.IsActive)
            {
                _DefaultServer.Stop();
            }
        }
        #endregion

        #region NdaHttpServer
        private NdaHttpServer(string ip, int port)
            : base(ip, port)
        {
        }
        #endregion

        #region HandleGetRequest
        public override void HandleGetRequest(HttpProcessor p)
        { 
            p.WriteSuccess();
            JObject resultObj = this.GetErrorString(new Exception("不支持get"));
            p._OutputStream.WriteLine(resultObj.ToString());
        }
        #endregion

        #region HandlePostRequest
        public override void HandlePostRequest(HttpProcessor p, StreamReader inputData)
        { 
            string httpUrl = p._HttpUrl;
            JObject resultObj = null;
            string baseUrl = "";
            string code = "";
            try
            {
                int questionMarkIndex = httpUrl.IndexOf('?');
                if (questionMarkIndex == -1)
                {
                    baseUrl = httpUrl;
                }
                else
                {
                    baseUrl = httpUrl.Substring(0, questionMarkIndex);
                }

                string methodName = Path.GetFileName(baseUrl);

                string requestInfo = inputData.ReadToEnd();

                switch (methodName.ToLower())
                {
                    case "createtask":
                        resultObj = this.CreateTask(requestInfo);
                        break;
                    case "deletetask":
                        resultObj = this.DeleteTask(requestInfo);
                        break;
                    case "gettasklist":
                        resultObj = this.GetTaskList(requestInfo);
                        break;
                    case "getdetailinfobytaskid":
                        resultObj = this.GetDetailInfoByTaskId(requestInfo);
                        break;
                    case "changetasklevel":
                        resultObj = this.ChangeTaskLevel(requestInfo);
                        break;
                    default:
                        throw new Exception("Unknown method. name = " + methodName);
                }
                code = "000";
            }
            catch (Exception ex)
            {
                resultObj = this.GetErrorString(ex); 
            }

            p._OutputStream.WriteLine(resultObj.ToString()); 
        }
        #endregion

        #region GetErrorString
        private JObject GetErrorString(Exception ex)
        {
            StringBuilder s = new StringBuilder();
            while (ex != null)
            {
                s.AppendLine(ex.Message);
                ex = ex.InnerException;
            }
            string errorInfo = s.ToString();
            JObject resultObj = new JObject();
            resultObj.Add("code", "001");
            resultObj.Add("errors", HttpUtility.UrlEncode(errorInfo));
            return resultObj;
        }
        #endregion

        #region CreateTask
        public JObject CreateTask(string requestInfo)
        {
            JObject requestJson = JObject.Parse(requestInfo);
            Task_Main task = new Task_Main();
            task.Description = requestJson.GetValue("description").ToString();
            task.Level = int.Parse(requestJson.GetValue("level").ToString());
            task.Name = requestJson.GetValue("name").ToString();
            task.SerialNumber = requestJson.GetValue("serialNumber").ToString();

            String taskId = TaskDataProcessor.CreateTask(task);
            return this.GetDetailInfoByTaskId(taskId);
        }
        #endregion

        #region DeleteTask
        public JObject DeleteTask(string requestInfo)
        {
            JObject requestJson = JObject.Parse(requestInfo);
            string taskId = requestJson.GetValue("taskId").ToString();
            TaskDataProcessor.DeleteTask(taskId);
            JObject resultObj = new JObject();
            resultObj.Add("code", "000");
            return resultObj;
        }
        #endregion

        #region GetTaskList
        public JObject GetTaskList(string requestInfo)
        {
            //可获取未执行、正在执行的、已经执行的任务列表
            //根据任务Id获取任务步骤信息
            JObject requestJson = JObject.Parse(requestInfo);
            int pageIndex = int.Parse(requestJson.GetValue("pageIndex").ToString());
            int onePageCount = int.Parse(requestJson.GetValue("onePageCount").ToString());
            List<Task_Main> tasks = TaskDataProcessor.GetTaskList(pageIndex, onePageCount);
            JObject resultObj = new JObject();
            resultObj.Add("code", "000");

            JArray taskObjects = new JArray();
            foreach (Task_Main task in tasks)
            {
                JObject taskObj = new JObject();
                taskObj.Add("createTime", task.CreateTime == null ? "" : ((DateTime)task.CreateTime).ToString("yyyy-MM-dd HH:mm:ss"));
                taskObj.Add("description", HttpUtility.UrlEncode(task.Description));
                taskObj.Add("createTime", task.EndTime == null ? "" : ((DateTime)task.EndTime).ToString("yyyy-MM-dd HH:mm:ss"));
                taskObj.Add("id", task.Id);
                taskObj.Add("level", task.Level);
                taskObj.Add("name", HttpUtility.UrlEncode(task.Name));
                taskObj.Add("serialNumber", HttpUtility.UrlEncode(task.SerialNumber));
                taskObj.Add("startTime", task.StartTime == null ? "" : ((DateTime)task.StartTime).ToString("yyyy-MM-dd HH:mm:ss"));
                taskObj.Add("statusType", task.StatusType.ToString());
                taskObjects.Add(taskObj);
            }
            resultObj.Add("tasks", taskObjects);

            return resultObj;
        }
        #endregion

        #region GetDetailInfoByTaskId
        public JObject GetDetailInfoByTaskId(string requestInfo)
        {
            //根据任务Id获取任务步骤信息
            JObject requestJson = JObject.Parse(requestInfo);
            string taskId = requestJson.GetValue("taskId").ToString();
            Dictionary<string, object> taskInfo = TaskDataProcessor.GetDetailInfoByTaskId(taskId);
            JObject resultObj = new JObject();
            resultObj.Add("code", "000");

            Task_Main task = (Task_Main)taskInfo["task"];
            JObject taskObj = new JObject();
            taskObj.Add("createTime", task.CreateTime == null ? "" : ((DateTime)task.CreateTime).ToString("yyyy-MM-dd HH:mm:ss"));
            taskObj.Add("description", HttpUtility.UrlEncode(task.Description));
            taskObj.Add("createTime", task.EndTime == null ? "" : ((DateTime)task.EndTime).ToString("yyyy-MM-dd HH:mm:ss"));
            taskObj.Add("id", task.Id);
            taskObj.Add("level", task.Level);
            taskObj.Add("name", HttpUtility.UrlEncode(task.Name));
            taskObj.Add("serialNumber", HttpUtility.UrlEncode(task.SerialNumber));
            taskObj.Add("startTime", task.StartTime == null ? "" : ((DateTime)task.StartTime).ToString("yyyy-MM-dd HH:mm:ss"));
            taskObj.Add("statusType", task.StatusType.ToString());
            resultObj.Add("task", taskObj);

            List<Task_Step> steps = (List<Task_Step>)taskInfo["steps"];
            JArray stepObjects = new JArray();
            foreach (Task_Step step in steps)
            {
                JObject stepObj = new JObject();
                stepObj.Add("endTime", step.EndTime == null ? "" : ((DateTime)step.EndTime).ToString("yyyy-MM-dd HH:mm:ss"));
                stepObj.Add("fatalErrorInfo", HttpUtility.UrlEncode(step.Message));
                stepObj.Add("id", step.Id);
                stepObj.Add("inputParameters", HttpUtility.UrlEncode(step.InputParameters));
                stepObj.Add("projectName", HttpUtility.UrlEncode(step.ProjectName));
                stepObj.Add("runIndex", step.RunIndex);
                stepObj.Add("startTime", step.StartTime == null ? "" : ((DateTime)step.StartTime).ToString("yyyy-MM-dd HH:mm:ss"));
                stepObj.Add("statusType", step.StatusType.ToString());
                stepObj.Add("taskId", step.TaskId);
                stepObjects.Add(stepObj);
            }
            resultObj.Add("steps", stepObjects);

            return resultObj;
        }
        #endregion

        #region ChangeTaskLevel
        public JObject ChangeTaskLevel(string requestInfo)
        {
            JObject requestJson = JObject.Parse(requestInfo);
            string taskId = requestJson.GetValue("taskId").ToString();
            int newLevel = int.Parse(requestJson.GetValue("newLevel").ToString());
            TaskDataProcessor.ChangeTaskLevel(taskId, newLevel);
            JObject resultObj = new JObject();
            resultObj.Add("code", "000");
            return resultObj;
        }
        #endregion
    }

}
