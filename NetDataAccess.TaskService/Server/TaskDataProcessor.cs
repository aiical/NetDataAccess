using NetDataAccess.Base.Common;
using NetDataTaskManager.Task;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;

namespace NetDataTaskManager.Manager
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

        #region 创建任务
        public static string CreateTask(Task_Main task)
        {
            TaskDataProcessor processor = new TaskDataProcessor();
            processor.Task = task;
            return (string)SqliteHelper.MainDbHelper.RunTransaction(new SqliteHelper.RunTransactionDelegate(processor.InsertTaskToDBTransaction));
        }

        private string InsertTaskToDBTransaction(SQLiteConnection dbConection)
        {
            Task_Main task = this.Task;
            if (task.AllSteps == null)
            {
                throw new Exception("Task steps can not be none!");
            }
            else
            { 
                string taskId = Guid.NewGuid().ToString();
                task.CreateTime = new DateTime();
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
                taskP2vs.Add("statusType", task.StatusType);
                taskP2vs.Add("description", task.Description);
                taskP2vs.Add("level", task.Level);
                taskP2vs.Add("serialNumber", task.SerialNumber);


                SQLiteCommand taskCmd = dbConection.CreateCommand();
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
                        projectName, 
                        taskId ,
                        statusType,
                        inputParamters, 
                        runIndex) 
                        values(
                        :id, 
                        :projectName, 
                        :taskId ,
                        :statusType,
                        :inputParamters, 
                        :runIndex)";
                    Dictionary<string, object> stepP2vs = new Dictionary<string, object>();
                    stepP2vs.Add("id", stepId);
                    stepP2vs.Add("projectName", step.ProjectName);
                    stepP2vs.Add("taskId", taskId);
                    stepP2vs.Add("statusType", step.StatusType);
                    stepP2vs.Add("inputParamters", step.InputParameters);
                    stepP2vs.Add("runIndex", step.RunIndex); 

                    SQLiteCommand stepCmd = dbConection.CreateCommand();
                    stepCmd.CommandText = addStepSql;
                    foreach (string pName in stepP2vs.Keys)
                    {
                        stepCmd.Parameters.AddWithValue(pName, stepP2vs[pName]);
                    }
                    stepCmd.ExecuteNonQuery();
                }

                return taskId;
            }
        }
        #endregion

        #region 获取任务列表
        public static List<Task_Main> GetTaskList(int pageIndex, int onePageCount)
        {
            int startIndex = (pageIndex - 1) * onePageCount;
            int endIndex = startIndex + onePageCount;
            string selectTaskSql = "select t.id as id, "
                + " t.name as name,"
                + " t.createtime as createtime,"
                + " t.starttime as starttime,"
                + " t.endtime as endtime,"
                + " t.statustype as statustype,"
                + " t.description as description,"
                + " t.level as level,"
                + " t.serialnumber as serialnumber"
                + " from task_main t"
                + " order by t.createtime desc"
                + " limit " + startIndex.ToString() + "," + endIndex.ToString();

            DataTable taskTable = SqliteHelper.MainDbHelper.GetDataTable(selectTaskSql, null);
            List<Task_Main> taskObjects = new List<Task_Main>();
            foreach (DataRow row in taskTable.Rows)
            {
                Task_Main taskObj = new Task_Main();
                taskObj.CreateTime = (DateTime)row["createtime"];
                taskObj.Description = (string)row["description"];
                taskObj.EndTime = (DateTime)row["endtime"];
                taskObj.Id = (string)row["id"];
                taskObj.Level = (int)row["level"];
                taskObj.Name = (string)row["name"];
                taskObj.SerialNumber = (string)row["serialnumber"];
                taskObj.StartTime = (DateTime)row["starttime"];
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
                + " t.starttime as starttime,"
                + " t.endtime as endtime,"
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
                taskObj.EndTime = (DateTime)taskRow["endtime"];
                taskObj.Id = (string)taskRow["id"];
                taskObj.Level = (int)taskRow["level"];
                taskObj.Name = (string)taskRow["name"];
                taskObj.SerialNumber = (string)taskRow["serialnumber"];
                taskObj.StartTime = (DateTime)taskRow["starttime"];
                string taskStatusTypeStr = (string)taskRow["statustype"];
                taskObj.StatusType = (TaskStatusType)Enum.Parse(typeof(TaskStatusType), taskStatusTypeStr);
                taskInfo.Add("task", taskObj);

                string selectStepSql = "select s.id as id, "
                    + " s.taskid as taskid,"
                    + " s.statustype as statustype,"
                    + " s.starttime as starttime,"
                    + " s.endtime as endtime,"
                    + " s.inputparameters as inputparameters,"
                    + " s.message as message,"
                    + " s.runindex as runindex,"
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
                    stepObj.EndTime = (DateTime)row["endtime"];
                    stepObj.Message = (string)row["message"];
                    stepObj.Id = (string)row["id"];
                    stepObj.InputParameters = (string)row["inputparameters"];
                    stepObj.ProjectName = (string)row["projectname"];
                    stepObj.RunIndex = (int)row["runindex"];
                    stepObj.StartTime = (DateTime)row["starttime"];
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

            string updateTaskSql = "update task_main set statustype = 'running' where id = :taskid";
            SQLiteCommand updateTaskCmd = new SQLiteCommand(conn);
            updateTaskCmd.CommandText = updateTaskSql;
            updateTaskCmd.Parameters.AddWithValue("taskid", step.TaskId);
            updateTaskCmd.ExecuteNonQuery();

            string updateStepSql = "update task_step set statustype = 'running' where id = :stepid";
            SQLiteCommand updateStepCmd = new SQLiteCommand(conn);
            updateStepCmd.CommandText = updateStepSql;
            updateStepCmd.Parameters.AddWithValue("taskid", Step.TaskId);
            updateStepCmd.ExecuteNonQuery();
            return true;
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
        }
        private void UpdateDataAfterEndStep(SQLiteConnection conn)
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
                string updateTaskSql = "update task_main set statustype = 'error' where id = :taskid";
                SQLiteCommand updateTaskCmd = new SQLiteCommand(conn);
                updateTaskCmd.CommandText = updateTaskSql; 
                updateTaskCmd.Parameters.AddWithValue("taskid",taskId);
                updateTaskCmd.ExecuteNonQuery(); 
            }
            else
            {

                string getWaitingCountStepSql = "select count(1) as stepCount from task_step s where s.taskid = :taskid and s.statustype = 'waiting'";
                SQLiteCommand getWaitingCountStepCmd = new SQLiteCommand(conn);
                getWaitingCountStepCmd.CommandText = getWaitingCountStepSql;
                getWaitingCountStepCmd.Parameters.AddWithValue("taskid", taskId);
                int stepCount = (int)getWaitingCountStepCmd.ExecuteScalar();

                if (stepCount == 0)
                {
                    string updateTaskSql = "update task_Main set statustype = 'succeed' where id = :taskid";
                    SQLiteCommand updateTaskCmd = new SQLiteCommand(conn);
                    updateTaskCmd.CommandText = updateTaskSql;
                    updateTaskCmd.Parameters.AddWithValue("taskid", taskId);
                    updateTaskCmd.ExecuteNonQuery();
                }
            }
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
                + " s.inputparameters as inputparameters,"
                + " s.message as message,"
                + " s.runindex as runindex,"
                + " s.projectname as projectname"
                + " from task_setp s"
                + " where s.taskid = :taskid and s.statustype = 'waiting'"
                + " order by s.runindex asc"
                + " limit 0, 0";
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
                stepObj.EndTime = (DateTime)row["endtime"];
                stepObj.Message = (string)row["message"];
                stepObj.Id = (string)row["id"];
                stepObj.InputParameters = (string)row["inputparameters"];
                stepObj.ProjectName = (string)row["projectname"];
                stepObj.RunIndex = (int)row["runindex"];
                stepObj.StartTime = (DateTime)row["starttime"];
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
            int endIndex = count - 1;

            string selectTaskSql = "";

            if (taskStatusType == TaskStatusType.Waiting)
            {
                //从从未有执行过step的task里选
                selectTaskSql = "select t.id as id"
                    + " from task_main t"
                    + " where t.statustype = 'waiting'"
                    + " order by t.level desc, t.createtime asc"
                    + " limit " + startIndex.ToString() + "," + endIndex.ToString();
            }
            else
            {
                //找到task状态为running，但是此task没有对应running的step
                selectTaskSql = "select t.id as id"
                    + " from task_main t"
                    + " where t.statustype = 'running'"
                    + " and not exists(select 1 from task_step s where s.taskid = t.id and s.statustype = 'running')"
                    + " order by t.level desc, t.createtime asc"
                    + " limit " + startIndex.ToString() + "," + endIndex.ToString();
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
