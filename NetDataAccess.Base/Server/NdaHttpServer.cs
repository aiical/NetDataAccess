using NetDataAccess.Base.Task;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Web;
using System.Xml;

namespace NetDataAccess.Base.Server
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

        #region 检测服务是否正常
        public static bool TestServer()
        {
            try
            {
                if (NdaHttpServer._DefaultServer == null)
                {
                    return false;
                }
                else
                {
                    WebClient wc = new WebClient();
                    string url = "http://" + NdaHttpServer._DefaultServer._IP + ":" + NdaHttpServer._DefaultServer._Port + "/test";
                    byte[] result = wc.UploadData(url, new byte[]{});
                    string resultStr = Encoding.UTF8.GetString(result).Trim();
                    return resultStr == "succeed";
                }
            }
            catch (Exception ex)
            {
                return false;
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
            XmlDocument resultDoc = this.GetErrorDoc(new Exception("不支持get"));
            p._OutputStream.WriteLine(resultDoc.DocumentElement.OuterXml);
        }
        #endregion

        #region HandlePostRequest modified by lixin 20170720
        public override void HandlePostRequest(HttpProcessor p, StreamReader inputData)
        {
            p.WriteSuccess();
            string httpUrl = p._HttpUrl;
            XmlDocument resultDoc = null;
            string baseUrl = ""; 
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

                string methodName = Path.GetFileName(baseUrl).ToLower();

                string requestInfo = inputData.ReadToEnd();
                
                //增加了test的命令判断
                if (methodName == "test")
                {
                    p._OutputStream.WriteLine("succeed");
                }
                else
                {
                    switch (methodName)
                    {
                        case "createtask":
                            resultDoc = this.CreateTask(requestInfo);
                            break;
                        case "deletetask":
                            resultDoc = this.DeleteTask(requestInfo);
                            break;
                        case "gettasklist":
                            resultDoc = this.GetTaskList(requestInfo);
                            break;
                        case "getdetailinfobytaskid":
                            resultDoc = this.GetDetailInfoByTaskId(requestInfo);
                            break;
                        case "changetasklevel":
                            resultDoc = this.ChangeTaskLevel(requestInfo);
                            break;
                        default:
                            throw new Exception("Unknown method. name = " + methodName);
                    }
                    p._OutputStream.WriteLine(resultDoc.DocumentElement.OuterXml);
                }
            }
            catch (Exception ex)
            {
                resultDoc = this.GetErrorDoc(ex);
                p._OutputStream.WriteLine(resultDoc.DocumentElement.OuterXml);
            }

        }
        #endregion

        #region GetErrorString
        private XmlDocument GetErrorDoc(Exception ex)
        {
            StringBuilder s = new StringBuilder();
            while (ex != null)
            {
                s.AppendLine(ex.Message);
                ex = ex.InnerException;
            }
            string errorInfo = s.ToString(); 

            XmlDocument resultDoc = new XmlDocument();
            resultDoc.LoadXml("<?xml version=\"1.0\" encoding=\"utf-8\" ?><result></result>");
            XmlElement resultRootNode = resultDoc.DocumentElement;
            resultRootNode.SetAttribute("code", "001");
            resultRootNode.SetAttribute("message", errorInfo);
            return resultDoc;
        }
        #endregion

        #region CreateTask
        public XmlDocument CreateTask(string requestInfo)
        {
            XmlDocument requestDoc = new XmlDocument();
            requestDoc.LoadXml(requestInfo);
            XmlElement rootNode = requestDoc.DocumentElement;
            Task_Main task = new Task_Main();
            task.Description = rootNode.GetAttribute("description");
            task.Level = int.Parse(rootNode.GetAttribute("level"));
            task.Name = rootNode.GetAttribute("name");
            task.SerialNumber = rootNode.GetAttribute("serialNumber");

            XmlNodeList allStepNodes = rootNode.SelectSingleNode("steps").SelectNodes("step");
            List<Task_Step> allSteps = new List<Task_Step>();
            foreach (XmlElement stepNode in allStepNodes)
            {
                Task_Step step = new Task_Step();
                step.GroupName = stepNode.GetAttribute("groupName");
                step.ProjectName = stepNode.GetAttribute("projectName");
                step.ListFilePath = stepNode.GetAttribute("listFilePath");
                step.InputDir = stepNode.GetAttribute("inputDir");
                step.MiddleDir = stepNode.GetAttribute("middleDir");
                step.OutputDir = stepNode.GetAttribute("outputDir");
                step.Parameters = stepNode.GetAttribute("parameters");
                step.RunIndex = int.Parse(stepNode.GetAttribute("runIndex"));
                allSteps.Add(step);
            }
            task.AllSteps = allSteps;

            String taskId = TaskDataProcessor.CreateTask(task);
            return this.GetDetailInfoByTaskIdValue(taskId);
        }
        #endregion

        #region DeleteTask
        public XmlDocument DeleteTask(string requestInfo)
        {
            XmlDocument requestDoc = new XmlDocument();
            requestDoc.LoadXml(requestInfo);
            XmlElement rootNode = requestDoc.DocumentElement;
            string taskId = rootNode.GetAttribute("taskId");
            TaskDataProcessor.DeleteTask(taskId);

            XmlDocument resultDoc = new XmlDocument();
            resultDoc.LoadXml("<?xml version=\"1.0\" encoding=\"utf-8\" ?><result></result>");
            XmlElement resultRootNode = resultDoc.DocumentElement; 
            resultRootNode.SetAttribute("code", "000");
            return resultDoc;
        }
        #endregion

        #region GetTaskList
        public XmlDocument GetTaskList(string requestInfo)
        {
            //可获取未执行、正在执行的、已经执行的任务列表
            //根据任务Id获取任务步骤信息
            XmlDocument requestDoc = new XmlDocument();
            requestDoc.LoadXml(requestInfo);
            XmlElement rootNode = requestDoc.DocumentElement;
            int pageIndex = int.Parse(rootNode.GetAttribute("pageIndex"));
            int onePageCount = int.Parse(rootNode.GetAttribute("onePageCount"));
            List<Task_Main> tasks = TaskDataProcessor.GetTaskList(pageIndex, onePageCount);

            XmlDocument resultDoc = new XmlDocument();
            resultDoc.LoadXml("<?xml version=\"1.0\" encoding=\"utf-8\" ?><result></result>");
            XmlElement resultRootNode = resultDoc.DocumentElement;
            resultRootNode.SetAttribute("code", "000");

            XmlElement taskListNode = resultDoc.CreateElement("tasks");
            resultRootNode.AppendChild(taskListNode);
            foreach (Task_Main task in tasks)
            {
                XmlElement taskNode = resultDoc.CreateElement("task");
                taskNode.SetAttribute("createTime", task.CreateTime == null ? "" : ((DateTime)task.CreateTime).ToString("yyyy-MM-dd HH:mm:ss"));
                taskNode.SetAttribute("description", task.Description);
                taskNode.SetAttribute("id", task.Id);
                taskNode.SetAttribute("level", task.Level.ToString());
                taskNode.SetAttribute("name",  task.Name);
                taskNode.SetAttribute("serialNumber", task.SerialNumber);
                taskNode.SetAttribute("statusType", task.StatusType.ToString());
                taskListNode.AppendChild(taskNode);
            }
            return resultDoc;
        }
        #endregion

        #region GetDetailInfoByTaskId
        public XmlDocument GetDetailInfoByTaskId(string requestInfo)
        {
            //根据任务Id获取任务步骤信息
            XmlDocument requestDoc = new XmlDocument();
            requestDoc.LoadXml(requestInfo);
            XmlElement rootNode = requestDoc.DocumentElement;
            string taskId = rootNode.GetAttribute("taskId");
            return this.GetDetailInfoByTaskIdValue(taskId);
        }
        #endregion

        #region GetDetailInfoByTaskId
        public XmlDocument GetDetailInfoByTaskIdValue(string taskId)
        {
            //根据任务Id获取任务步骤信息 
            Dictionary<string, object> taskInfo = TaskDataProcessor.GetDetailInfoByTaskId(taskId);

            XmlDocument resultDoc = new XmlDocument();
            resultDoc.LoadXml("<?xml version=\"1.0\" encoding=\"utf-8\" ?><result></result>");
            XmlElement resultRootNode = resultDoc.DocumentElement;
            resultRootNode.SetAttribute("code", "000");
             
            Task_Main task = (Task_Main)taskInfo["task"];
            resultRootNode.SetAttribute("createTime", task.CreateTime == null ? "" : ((DateTime)task.CreateTime).ToString("yyyy-MM-dd HH:mm:ss"));
            resultRootNode.SetAttribute("description", task.Description);
            resultRootNode.SetAttribute("id", task.Id);
            resultRootNode.SetAttribute("level", task.Level.ToString());
            resultRootNode.SetAttribute("name", task.Name);
            resultRootNode.SetAttribute("serialNumber", task.SerialNumber);
            resultRootNode.SetAttribute("statusType", task.StatusType.ToString()); 

            List<Task_Step> steps = (List<Task_Step>)taskInfo["steps"];
            XmlElement stepListNode = resultDoc.CreateElement("steps");
            resultRootNode.AppendChild(stepListNode);
            foreach (Task_Step step in steps)
            {
                XmlElement stepNode = resultDoc.CreateElement("step");
                stepNode.SetAttribute("endTime", step.EndTime == null ? "" : ((DateTime)step.EndTime).ToString("yyyy-MM-dd HH:mm:ss"));
                stepNode.SetAttribute("message", step.Message);
                stepNode.SetAttribute("id", step.Id);
                stepNode.SetAttribute("listFilePath",step.ListFilePath);
                stepNode.SetAttribute("inputDir",step.InputDir);
                stepNode.SetAttribute("middleDir", step.MiddleDir);
                stepNode.SetAttribute("outputDir", step.OutputDir);
                stepNode.SetAttribute("parameters", step.Parameters);
                stepNode.SetAttribute("groupName", step.GroupName);
                stepNode.SetAttribute("projectName", step.ProjectName);
                stepNode.SetAttribute("runIndex", step.RunIndex.ToString());
                stepNode.SetAttribute("startTime", step.StartTime == null ? "" : ((DateTime)step.StartTime).ToString("yyyy-MM-dd HH:mm:ss"));
                stepNode.SetAttribute("statusType", step.StatusType.ToString());
                stepNode.SetAttribute("taskId", step.TaskId);
                stepListNode.AppendChild(stepNode);
            }

            return resultDoc;
        }
        #endregion

        #region ChangeTaskLevel
        public XmlDocument ChangeTaskLevel(string requestInfo)
        {
            XmlDocument requestDoc = new XmlDocument();
            requestDoc.LoadXml(requestInfo);
            XmlElement rootNode = requestDoc.DocumentElement;
            string taskId = rootNode.GetAttribute("taskId");

            int newLevel = int.Parse(rootNode.GetAttribute("newLevel"));
            TaskDataProcessor.ChangeTaskLevel(taskId, newLevel);

            XmlDocument resultDoc = new XmlDocument();
            resultDoc.LoadXml("<?xml version=\"1.0\" encoding=\"utf-8\" ?><result></result>");
            XmlElement resultRootNode = resultDoc.DocumentElement;
            resultRootNode.SetAttribute("code", "000");
            return resultDoc;
        }
        #endregion
    }

}
