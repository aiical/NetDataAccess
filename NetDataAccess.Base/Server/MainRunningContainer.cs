using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.Server
{
    public interface IMainRunningContainer
    {
        void InvokeRunTask(string groupName, string projectName, string listFilePath, string inputDir, string middleDir, string outputDir, string parameter, string stepId, bool autoRun, bool popPrompt);

        void InvokeCloseTaskUI(string taskId);
    }
}
