using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Main
{
    public interface MainRunningContainer
    {
        void InvokeRunTask(string groupName, string projectName, string parameter, string taskId);

        void InvokeCloseTaskUI(string taskId);
    }
}
