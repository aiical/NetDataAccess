using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.Task
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

        #region 分组名称
        private string _GroupName;
        public string GroupName
        {
            get
            {
                return _GroupName;
            }
            set
            {
                _GroupName = value;
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

        #region 下载列表文件地址
        private string _ListFilePath;
        public string ListFilePath
        {
            get
            {
                return _ListFilePath;
            }
            set
            {
                _ListFilePath = value;
            }
        }
        #endregion

        #region 输出目录
        private string _OutputDir;
        public string OutputDir
        {
            get
            {
                return _OutputDir;
            }
            set
            {
                _OutputDir = value;
            }
        }
        #endregion

        #region 输入目录
        private string _InputDir;
        public string InputDir
        {
            get
            {
                return _InputDir;
            }
            set
            {
                _InputDir = value;
            }
        }
        #endregion

        #region 中间目录
        private string _MiddleDir;
        public string MiddleDir
        {
            get
            {
                return _MiddleDir;
            }
            set
            {
                _MiddleDir = value;
            }
        }
        #endregion

        #region 输入参数
        private string _Parameters;
        public string Parameters
        {
            get
            {
                return _Parameters;
            }
            set
            {
                _Parameters = value;
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
