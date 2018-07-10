using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataTaskManager.Task
{
    public class Task_Step
    { 
        #region Id
        private string _Id;
        public string Id
        {
            get
            {
                return _Id;
            }
            set
            {
                _Id = value;
            }
        }
        #endregion

        #region TaskId
        private string _TaskId;
        public string TaskId
        {
            get
            {
                return _TaskId;
            }
            set
            {
                _TaskId = value;
            }
        }
        #endregion

        #region 项目名称
        private string _ProjectName;
        public string ProjectName
        {
            get
            {
                return _ProjectName;
            }
            set
            {
                _ProjectName = value;
            }
        }
        #endregion

        #region 状态
        private TaskStatusType _StatusType = TaskStatusType.Waiting;
        public TaskStatusType StatusType
        {
            get
            {
                return _StatusType;
            }
            set
            {
                _StatusType = value;
            }
        }
        #endregion

        #region 开始时间
        private Nullable<DateTime> _StartTime;
        public Nullable<DateTime> StartTime
        {
            get
            {
                return _StartTime;
            }
            set
            {
                _StartTime = value;
            }
        }
        #endregion

        #region 结束时间
        private Nullable<DateTime> _EndTime;
        public Nullable<DateTime> EndTime
        {
            get
            {
                return _EndTime;
            }
            set
            {
                _EndTime = value;
            }
        }
        #endregion

        #region 输入参数
        private string _InputParameters;
        public string InputParameters
        {
            get
            {
                return _InputParameters;
            }
            set
            {
                _InputParameters = value;
            }
        }
        #endregion

        #region 信息
        private string _Message;
        public string Message
        {
            get
            {
                return _Message;
            }
            set
            {
                _Message = value;
            }
        }
        #endregion

        #region 执行顺序
        private int _RunIndex;
        public int RunIndex
        {
            get
            {
                return _RunIndex;
            }
            set
            {
                _RunIndex = value;
            }
        }
        #endregion
    }
}
