using NetDataAccess.Base.Common;
using NetDataAccess.Base.Task;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.Server
{
    public class TaskDataProcessor
    {
        #region 构造函数
        private TaskDataProcessor()
        { 
        }
        #endregion

        #region Task
        private Task_Main _Task = null;
        private Task_Main Task
        {
            get
            {
                return _Task;
            }
            set
            {
                _Task = value;
            }
        }
        #endregion

        #region Step
        private Task_Step _Step = null;
        private Task_Step Step
        {
            get
            {
                return _Step;
            }
            set
            {
                _Step = value;
            }
        }
        #endregion

        #region 删除任务
        public static bool DeleteTask(string taskId)
        {
            TaskDataProcessor processor = new TaskDataProcessor();
            processor.Task = new Task_Main();
            processor.Task.Id = taskId;
            return (bool)SqliteHelper.MainDbHelper.RunTransaction(processor.DeleteTaskTransaction);
        }

        private object DeleteTaskTransaction(SQLiteConnection conn)
        {
            Task_Main task = this.Task;

            string selectTaskSql = "select t.id as id, t.statustype as statusType, t.serialNumber as serialNumber from task_Main t where t.id = :id for update";
            SQLiteCommand selectTaskCmd = new SQLiteCommand(conn);
            selectTaskCmd.CommandText = selectTaskSql; 
            selectTaskCmd.Parameters.AddWithValue("id", task.Id);
            SQLiteDataAdapter dao = new SQLiteDataAdapter(selectTaskCmd);
            DataTable taskTable = new DataTable();
            dao.Fill(taskTable);

            if (taskTable == null || taskTable.Rows.Count == 0)
            {
                throw new Exception("没有找到对应的任务. TaskId = " + task.Id);
            }
            else
            {
                DataRow taskRow = taskTable.Rows[0];
                String statusTypeStr = taskRow["statusType"].ToString();
                TaskStatusType statusType = (TaskStatusType)Enum.Parse(typeof(TaskStatusType), statusTypeStr);
                if (statusType == TaskStatusType.Deleted)
                {
                    return true;
                }
                else if (statusType == TaskStatusType.Error)
                {
                    throw new Exception("Can not delete a error task. TaskId = " + task.Id);
                }
                else if (statusType == TaskStatusType.Running)
                {
                    //可允许删除，但是需要和爬取工具交互，停止任务
                    //目前先限定住，不允许删除
                    throw new Exception("Can not delete a running task. TaskId = " + task.Id);
                }
                else if (statusType == TaskStatusType.Succeed)
                {
                    throw new Exception("Can not delete a succeed task. TaskId = " + task.Id);
                }
                else if (statusType == TaskStatusType.Waiting)
                {
                    string updateTaskSql = "update task_Main set statustype = :statusType where id = :id";
                    SQLiteCommand updateTaskCmd = new SQLiteCommand(conn);
                    updateTaskCmd.CommandText = updateTaskSql;
                    updateTaskCmd.Parameters.AddWithValue("statusType", TaskStatusType.Deleted.ToString());
                    updateTaskCmd.ExecuteNonQuery();
                    return true;
                }
                else
                {
                    throw new Exception("Can not delete a " + statusType.ToString() + " task. TaskId = " + task.Id);
                }
            }
        }
        #endregion

        #region CheckExistBySerialNumber
        public String CheckExistBySerialNumber(string serialNumber, SQLiteConnection conn)
        {
            string getTaskIdSql = "select t.id as taskId from task_main t where t.serialnumber = :serialnumber";
            SQLiteCommand getTaskIdCmd = new SQLiteCommand(conn);
            getTaskIdCmd.CommandText = getTaskIdSql;
            getTaskIdCmd.Parameters.AddWithValue("serialnumber", serialNumber);
            String taskId = (String)getTaskIdCmd.ExecuteScalar();
            return taskId;
        }
        #endregion

        #region 创建任务
        public static string CreateTask(Task_Main task)
        {
            TaskDataProcessor processor = new TaskDataProcessor();
            processor.Task = task;
            return (string)SqliteHelper.MainDbHelper.RunTransaction(new SqliteHelper.RunTransactionDelegate(processor.InsertTaskToDBTransaction));
        }

        private string InsertTaskToDBTransaction(SQLiteConnection conn)
        {
            Task_Main task = this.Task;
            if (task.AllSteps == null)
            {
                throw new Exception("Task steps can not be none!");
            }
            else 
            {
                string taskId = this.CheckExistBySerialNumber(task.SerialNumber, conn);

                if (taskId == null || taskId.Length == 0)
                {
                    taskId = Guid.NewGuid().ToString();
                    task.CreateTime = DateTime.Now;
                    task.StatusType = TaskStatusType.Waiting;
                    string addTaskSql = @"insert into task_Main(
                    id, 
                    name, 
                    createTime ,
                    statusType,
                    description, 
                    level, 
                    serialNumber) 
                    values(
                    :id, 
                    :name, 
                    :createTime ,
                    :statusType,
                    :description, 
                    :level, 
                    :serialNumber)";
                    Dictionary<string, object> taskP2vs = new Dictionary<string, object>();
                    taskP2vs.Add("id", taskId);
                    taskP2vs.Add("name", task.Name);
                    taskP2vs.Add("createTime", task.CreateTime);
                    taskP2vs.Add("statusType", task.StatusType.ToString());
                    taskP2vs.Add("description", task.Description);
                    taskP2vs.Add("level", task.Level);
                    taskP2vs.Add("serialNumber", task.SerialNumber);

                    SQLiteCommand taskCmd = conn.CreateCommand();
                    taskCmd.CommandText = addTaskSql;
                    foreach (string pName in taskP2vs.Keys)
                    {
                        taskCmd.Parameters.AddWithValue(pName, taskP2vs[pName]);
                    }
                    taskCmd.ExecuteNonQuery();

                    //添加step
                    foreach (Task_Step step in task.AllSteps)
                    {
                        string stepId = Guid.NewGuid().ToString();
                        step.StatusType = TaskStatusType.Waiting;
                        string addStepSql = @"insert into task_Step(
                        id, 
                        groupName, 
                        projectName, 
                        listFilePath, 
                        inputDir, 
                        middleDir,
                        outputDir, 
                        taskId ,
                        statusType,
                        parameters, 
                        runIndex) 
                        values(
                        :id, 
                        :groupName, 
                        :projectName, 
                        :listFilePath, 
                        :inputDir, 
                        :middleDir, 
                        :outputDir, 
                        :taskId ,
                        :statusType,
                        :parameters, 
                        :runIndex)";
                        Dictionary<string, object> stepP2vs = new Dictionary<string, object>();
                        stepP2vs.Add("id", stepId);
                        stepP2vs.Add("groupName", step.GroupName);
                        stepP2vs.Add("projectName", step.ProjectName);
                        stepP2vs.Add("taskId", taskId);
                        stepP2vs.Add("statusType", step.StatusType.ToString());
                        stepP2vs.Add("listFilePath", step.ListFilePath);
                        stepP2vs.Add("inputDir", step.InputDir);
                        stepP2vs.Add("middleDir", step.MiddleDir);
                        stepP2vs.Add("outputDir", step.OutputDir);
                        stepP2vs.Add("parameters", step.Parameters);
                        stepP2vs.Add("runIndex", step.RunIndex);

                        SQLiteCommand stepCmd = conn.CreateCommand();
                        stepCmd.CommandText = addStepSql;
                        foreach (string pName in stepP2vs.Keys)
                        {
                            stepCmd.Parameters.AddWithValue(pName, stepP2vs[pName]);
                        }
                        stepCmd.ExecuteNonQuery();
                    }
                }
                return taskId;
            }
        }
        #endregion

        #region 获取任务列表
        public static List<Task_Main> GetTaskList(int pageIndex, int onePageCount)
        {
            int startIndex = (pageIndex - 1) * onePageCount; 
            string selectTaskSql = "select t.id as id, "
                + " t.name as name,"
                + " t.createtime as createtime," 
                + " t.statustype as statustype,"
                + " t.description as description,"
                + " t.level as level,"
                + " t.serialnumber as serialnumber"
                + " from task_main t"
                + " order by t.createtime desc"
                + " limit " + onePageCount.ToString() + " offset " + startIndex.ToString();

            DataTable taskTable = SqliteHelper.MainDbHelper.GetDataTable(selectTaskSql, null);
            List<Task_Main> taskObjects = new List<Task_Main>();
            foreach (DataRow row in taskTable.Rows)
            {
                Task_Main taskObj = new Task_Main();
                taskObj.CreateTime = (DateTime)row["createtime"];
                taskObj.Description = (string)row["description"]; 
                taskObj.Id = (string)row["id"];
                taskObj.Level = (int)row["level"];
                taskObj.Name = (string)row["name"];
                taskObj.SerialNumber = (string)row["serialnumber"]; 
                string statusTypeStr = (string)row["statustype"];
                taskObj.StatusType = (TaskStatusType)Enum.Parse(typeof(TaskStatusType), statusTypeStr);
                taskObjects.Add(taskObj);
            }
            return taskObjects;
        }
        #endregion

        #region 获取任务步骤信息
        public static Dictionary<string,object> GetDetailInfoByTaskId(string taskId)
        {
            Dictionary<string, object> taskInfo = new Dictionary<string, object>();

            string selectTaskSql = "select t.id as id, "
                + " t.name as name,"
                + " t.createtime as createtime,"  
                + " t.statustype as statustype,"
                + " t.description as description,"
                + " t.level as level,"
                + " t.serialnumber as serialnumber"
                + " from task_main t"
                + " where t.id = :id"
                + " order by t.createtime desc";

            Dictionary<string, object> taskP2vs = new Dictionary<string, object>();
            taskP2vs.Add("id", taskId);

            DataTable taskTable = SqliteHelper.MainDbHelper.GetDataTable(selectTaskSql, taskP2vs);
            if (taskTable.Rows.Count == 0)
            {
                throw new Exception("None task. taskId = " + taskId);
            }
            else
            {
                DataRow taskRow = taskTable.Rows[0];
                Task_Main taskObj = new Task_Main();
                taskObj.CreateTime = (DateTime)taskRow["createtime"];
                taskObj.Description = (string)taskRow["description"]; 
                taskObj.Id = (string)taskRow["id"];
                taskObj.Level = (int)taskRow["level"];
                taskObj.Name = (string)taskRow["name"];
                taskObj.SerialNumber = (string)taskRow["serialnumber"]; 
                string taskStatusTypeStr = (string)taskRow["statustype"];
                taskObj.StatusType = (TaskStatusType)Enum.Parse(typeof(TaskStatusType), taskStatusTypeStr);
                taskInfo.Add("task", taskObj);

                string selectStepSql = "select s.id as id, "
                    + " s.taskid as taskid,"
                    + " s.statustype as statustype,"
                    + " s.starttime as starttime,"
                    + " s.endtime as endtime,"
                    + " s.listfilepath as listfilepath,"
                    + " s.inputdir as inputdir,"
                    + " s.middledir as middledir,"
                    + " s.outputdir as outputdir,"
                    + " s.parameters as parameters,"
                    + " s.message as message,"
                    + " s.runindex as runindex,"
                    + " s.groupname as groupname,"
                    + " s.projectname as projectname"
                    + " from task_step s"
                    + " where s.taskid = :taskid"
                    + " order by s.runindex desc";

                Dictionary<string, object> stepP2vs = new Dictionary<string, object>();
                stepP2vs.Add("taskid", taskId);

                DataTable stepTable = SqliteHelper.MainDbHelper.GetDataTable(selectStepSql, stepP2vs);
                List<Task_Step> stepObjects = new List<Task_Step>();
                foreach (DataRow row in stepTable.Rows)
                {
                    Task_Step stepObj = new Task_Step();
                    stepObj.EndTime = row["endtime"] == null || row["endtime"] == DBNull.Value ? null : (Nullable<DateTime>)(DateTime)row["endtime"];
                    stepObj.Message = row["message"] == null || row["message"] == DBNull.Value ? "" : (string)row["message"];
                    stepObj.Id = (string)row["id"];
                    stepObj.ListFilePath = (string)row["listfilepath"];
                    stepObj.InputDir = (string)row["inputdir"];
                    stepObj.MiddleDir = (string)row["middledir"];
                    stepObj.OutputDir = (string)row["outputdir"];
                    stepObj.Parameters = (string)row["parameters"];
                    stepObj.GroupName = (string)row["groupname"];
                    stepObj.ProjectName = (string)row["projectname"];
                    stepObj.RunIndex = int.Parse(row["runindex"].ToString());
                    stepObj.StartTime = row["starttime"] == null || row["starttime"] == DBNull.Value ? null : (Nullable<DateTime>)(DateTime)row["starttime"];
                    stepObj.TaskId = (string)row["taskid"];
                    string statusTypeStr = (string)row["statustype"];
                    stepObj.StatusType = (TaskStatusType)Enum.Parse(typeof(TaskStatusType), statusTypeStr);
                    stepObjects.Add(stepObj);
                }
                taskInfo.Add("steps", stepObjects);

                return taskInfo;
            }
        }
        #endregion

        #region 修改任务优先级
        public static bool ChangeTaskLevel(string taskId, int newLevel)
        {
            TaskDataProcessor processor = new TaskDataProcessor();
            processor.Task = new Task_Main();
            processor.Task.Id = taskId;
            processor.Task.Level = newLevel;
            return (bool)SqliteHelper.MainDbHelper.RunTransaction(processor.ChangeTaskLevelTransaction);
        }

        private object ChangeTaskLevelTransaction(SQLiteConnection conn)
        {
            Task_Main task = this.Task;

            string selectTaskSql = "select t.id as id, t.statustype as statusType, t.serialNumber as serialNumber from task_Main t where t.id = :id for update";
            SQLiteCommand selectTaskCmd = new SQLiteCommand(conn);
            selectTaskCmd.CommandText = selectTaskSql;
            selectTaskCmd.Parameters.AddWithValue("id", task.Id);
            SQLiteDataAdapter dao = new SQLiteDataAdapter(selectTaskCmd);
            DataTable taskTable = new DataTable();
            dao.Fill(taskTable);

            if (taskTable == null || taskTable.Rows.Count == 0)
            {
                throw new Exception("没有找到对应的任务. TaskId = " + task.Id);
            }
            else
            {
                DataRow taskRow = taskTable.Rows[0];
                String statusTypeStr = taskRow["statusType"].ToString();
                TaskStatusType statusType = (TaskStatusType)Enum.Parse(typeof(TaskStatusType), statusTypeStr);
                if (statusType == TaskStatusType.Deleted)
                {
                    return true;
                }
                else if (statusType == TaskStatusType.Error)
                {
                    throw new Exception("Can not change a error task level. TaskId = " + task.Id);
                }
                else if (statusType == TaskStatusType.Running)
                {
                    throw new Exception("Can not change a running task level. TaskId = " + task.Id);
                }
                else if (statusType == TaskStatusType.Succeed)
                {
                    throw new Exception("Can not change a succeed task level. TaskId = " + task.Id);
                }
                else if (statusType == TaskStatusType.Waiting)
                {
                    string updateTaskSql = "update task_Main set level = :level where id = :id";
                    SQLiteCommand updateTaskCmd = new SQLiteCommand(conn);
                    updateTaskCmd.CommandText = updateTaskSql;
                    updateTaskCmd.Parameters.AddWithValue("level", task.Level);
                    updateTaskCmd.ExecuteNonQuery();
                    return true;
                }
                else
                {
                    throw new Exception("Can not change a " + statusType.ToString() + " task level. TaskId = " + task.Id);
                }
            }
        }
        #endregion

        #region 开始执行Step时，修改数据库信息，先修改数据，再开始执行爬取任务
        public static void UpdateDataAfterBeginStep(string stepId, string taskId)
        {
            TaskDataProcessor processor = new TaskDataProcessor(); 
            processor.Step = new Task_Step();
            processor.Step.Id = stepId;
            processor.Step.TaskId = taskId;
            processor.Step.StatusType = TaskStatusType.Running;
            SqliteHelper.MainDbHelper.RunTransaction(processor.UpdateDataAfterBeginStep);
        }

        private object UpdateDataAfterBeginStep(SQLiteConnection conn)
        {
            Task_Step step = this.Step;
            DateTime currentTime = DateTime.Now;

            string updateTaskSql = "update task_main set statustype = 'Running' where id = :taskid";
            SQLiteCommand updateTaskCmd = new SQLiteCommand(conn);
            updateTaskCmd.CommandText = updateTaskSql; 
            updateTaskCmd.Parameters.AddWithValue("taskid", step.TaskId);
            updateTaskCmd.ExecuteNonQuery();

            string updateStepSql = "update task_step set statustype = 'Running', starttime = :starttime  where id = :stepid";
            SQLiteCommand updateStepCmd = new SQLiteCommand(conn);
            updateStepCmd.CommandText = updateStepSql;
            updateStepCmd.Parameters.AddWithValue("starttime", currentTime);
            updateStepCmd.Parameters.AddWithValue("stepid", Step.Id);
            updateStepCmd.ExecuteNonQuery();
            return true;
        }
        #endregion

        #region 启动时调用，重置Running的step为waiting状态
        public static void ResetRunningSteps()
        {
            string updateStepSql = "update task_step set statustype = :statustype where statustype = :oldstatustype";
            Dictionary<string, object> p2vs = new Dictionary<string, object>();
            p2vs.Add("statustype", TaskStatusType.Waiting.ToString());
            p2vs.Add("oldstatustype", TaskStatusType.Running.ToString());
            SqliteHelper.MainDbHelper.ExecuteSql(updateStepSql, p2vs);
        }
        #endregion

        #region Step运行完后，更新数据
        public static void UpdateDataAfterEndStep(string stepId, string taskId, bool succeed, string message)
        {
            TaskDataProcessor processor = new TaskDataProcessor();
            processor.Step = new Task_Step();
            processor.Step.Id = stepId;
            processor.Step.TaskId = taskId;
            processor.Step.StatusType = succeed ? TaskStatusType.Succeed : TaskStatusType.Error;
            processor.Step.Message = message;
            SqliteHelper.MainDbHelper.RunTransaction(processor.UpdateDataAfterEndStep);
        }
        private object UpdateDataAfterEndStep(SQLiteConnection conn)
        {
            Task_Step step = this.Step;

            string taskId = step.TaskId;
            string stepId = step.Id;

            string updateStepSql = "update task_step set statustype = :statustype, message = :message where id = :stepid";
            SQLiteCommand updateStepCmd = new SQLiteCommand(conn);
            updateStepCmd.CommandText = updateStepSql;
            updateStepCmd.Parameters.AddWithValue("statustype", step.StatusType.ToString());
            updateStepCmd.Parameters.AddWithValue("stepid", stepId);
            updateStepCmd.Parameters.AddWithValue("message", step.Message);
            updateStepCmd.ExecuteNonQuery();

            if (step.StatusType == TaskStatusType.Error)
            {
                string updateTaskSql = "update task_main set statustype = 'Error' where id = :taskid";
                SQLiteCommand updateTaskCmd = new SQLiteCommand(conn);
                updateTaskCmd.CommandText = updateTaskSql; 
                updateTaskCmd.Parameters.AddWithValue("taskid",taskId);
                updateTaskCmd.ExecuteNonQuery();
            }
            else
            {

                string getWaitingCountStepSql = "select count(1) as stepCount from task_step s where s.taskid = :taskid and s.statustype = 'Waiting'";
                SQLiteCommand getWaitingCountStepCmd = new SQLiteCommand(conn);
                getWaitingCountStepCmd.CommandText = getWaitingCountStepSql;
                getWaitingCountStepCmd.Parameters.AddWithValue("taskid", taskId);
                int stepCount = int.Parse(getWaitingCountStepCmd.ExecuteScalar().ToString());

                if (stepCount == 0)
                {
                    string updateTaskSql = "update task_Main set statustype = 'Succeed' where id = :taskid";
                    SQLiteCommand updateTaskCmd = new SQLiteCommand(conn);
                    updateTaskCmd.CommandText = updateTaskSql;
                    updateTaskCmd.Parameters.AddWithValue("taskid", taskId);
                    updateTaskCmd.ExecuteNonQuery();
                }
            }
            return true;
        }
        #endregion

        #region 获取该任务里该执行的step
        private Task_Step GetNextStepInTask(string taskId)
        {
            string getNextStepIdSql = "select s.id as id, "
                + " s.taskid as taskid,"
                + " s.statustype as statustype,"
                + " s.starttime as starttime,"
                + " s.endtime as endtime,"
                + " s.listfilepath as listfilepath,"
                + " s.inputdir as inputdir,"
                + " s.middledir as middledir,"
                + " s.outputdir as outputdir,"
                + " s.parameters as parameters,"
                + " s.message as message,"
                + " s.runindex as runindex,"
                + " s.groupname as groupname,"
                + " s.projectname as projectname"
                + " from task_step s"
                + " where s.taskid = :taskid and s.statustype = 'Waiting'"
                + " order by s.runindex asc"
                + " limit 1";
            Dictionary<string, object> p2vs = new Dictionary<string, object>();
            p2vs.Add("taskid", taskId);
            DataTable stepTable = SqliteHelper.MainDbHelper.GetDataTable(getNextStepIdSql, p2vs);
            if (stepTable.Rows.Count == 0)
            {
                return null;
            }
            else
            {
                DataRow row = stepTable.Rows[0];
                Task_Step stepObj = new Task_Step();
                stepObj.EndTime = CommonUtil.IsNullOrDBNul(row["endtime"]) ? null : (Nullable<DateTime>)row["endtime"];
                stepObj.Message = CommonUtil.IsNullOrDBNul(row["message"]) ? null : (string)row["message"];
                stepObj.Id =  (string)row["id"];
                stepObj.ListFilePath = CommonUtil.IsNullOrDBNul(row["listfilepath"]) ? null : (string)row["listfilepath"];
                stepObj.InputDir = CommonUtil.IsNullOrDBNul(row["inputdir"]) ? null : (string)row["inputdir"];
                stepObj.MiddleDir = CommonUtil.IsNullOrDBNul(row["middledir"]) ? null : (string)row["middledir"];
                stepObj.OutputDir = CommonUtil.IsNullOrDBNul(row["outputdir"]) ? null : (string)row["outputdir"];
                stepObj.Parameters = CommonUtil.IsNullOrDBNul(row["parameters"]) ? null : (string)row["parameters"];
                stepObj.GroupName = (string)row["groupname"];
                stepObj.ProjectName = (string)row["projectname"];
                stepObj.RunIndex = int.Parse(row["runindex"].ToString());
                stepObj.StartTime = CommonUtil.IsNullOrDBNul(row["starttime"]) ? null : (Nullable<DateTime>)row["starttime"];
                stepObj.TaskId = (string)row["taskid"];
                string statusTypeStr = (string)row["statustype"];
                stepObj.StatusType = (TaskStatusType)Enum.Parse(typeof(TaskStatusType), statusTypeStr);
                return stepObj;
            }
        }
        #endregion

        #region 获取需要执行的N个任务步骤
        public static List<Task_Step> GetWaitingStepsInRunningTasks(int count)
        {
            TaskDataProcessor processor = new TaskDataProcessor();
            return processor.GetWaitingStepsInTasks(count, TaskStatusType.Running);
        } 
        public static List<Task_Step> GetWaitingStepsInWaitingTasks(int count)
        {
            TaskDataProcessor processor = new TaskDataProcessor();
            return processor.GetWaitingStepsInTasks(count, TaskStatusType.Waiting);
        } 
        private List<Task_Step> GetWaitingStepsInTasks(int count, TaskStatusType taskStatusType)
        {
            int startIndex = 0; 

            string selectTaskSql = "";

            if (taskStatusType == TaskStatusType.Waiting)
            {
                //从从未有执行过step的task里选
                selectTaskSql = "select t.id as id"
                    + " from task_main t"
                    + " where t.statustype = 'Waiting'"
                    + " order by t.level desc, t.createtime asc"
                    + " limit " + count.ToString() + " offset " + startIndex.ToString();
            }
            else
            {
                //找到task状态为running，但是此task没有对应running的step
                selectTaskSql = "select t.id as id"
                    + " from task_main t"
                    + " where t.statustype = 'Running'"
                    + " and not exists(select 1 from task_step s where s.taskid = t.id and s.statustype = 'Running')"
                    + " order by t.level desc, t.createtime asc"
                    + " limit " + count.ToString() + " offset " + startIndex.ToString();
            }

            DataTable taskTable = SqliteHelper.MainDbHelper.GetDataTable(selectTaskSql, null);

            List<Task_Step> stepObjects = new List<Task_Step>();
            foreach (DataRow row in taskTable.Rows)
            {
                string taskId = (string)row["id"];
                Task_Step step = this.GetNextStepInTask(taskId);
                if (step != null)
                {
                    stepObjects.Add(step);
                }
            }
            return stepObjects;
        }
        #endregion
    }
}
