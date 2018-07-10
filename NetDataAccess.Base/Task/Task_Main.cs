using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.Task
{
    public class Task_Main
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

        #region Name
        private string _Name;
        public string Name
        {
            get
            {
                return _Name;
            }
            set
            {
                _Name = value;
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

        #region 创建时间
        private Nullable<DateTime> _CreateTime;
        public Nullable<DateTime> CreateTime 
        {
            get
            {
                return _CreateTime;
            }
            set
            {
                _CreateTime = value;
            }
        }
        #endregion 

        #region 描述
        private string _Description;
        public string Description
        {
            get
            {
                return _Description;
            }
            set
            {
                _Description = value;
            }
        }
        #endregion

        #region 抓取步骤
        private List<Task_Step> _AllSteps = null;
        public List<Task_Step> AllSteps
        {
            get
            {
                return _AllSteps;
            }
            set
            {
                _AllSteps = value;
            }
        }
        #endregion

        #region 优先级
        private int _Level = 0;
        public int Level
        {
            get
            {
                return _Level;
            }
            set
            {
                _Level = value;
            }
        }
        #endregion


        #region 流水号
        private string _SerialNumber;
        public string SerialNumber
        {
            get
            {
                return _SerialNumber;
            }
            set
            {
                _SerialNumber = value;
            }
        }
        #endregion
    }
}
