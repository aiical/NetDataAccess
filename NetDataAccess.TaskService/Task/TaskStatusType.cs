using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataTaskManager.Task
{
    public enum TaskStatusType
    {
        Waiting,
        Running,
        Succeed,
        Error,
        Deleted
    }
}
