using NetDataAccess.Base.Proxy;
using NetDataAccess.Base.Web;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace NetDataAccess.Base.Common
{
    public class FileRequest
    {
        #region 通过webrequest获取File
        public byte[] GetFileByRequest(string pageUrl , decimal intervalAfterLoaded, int timeout)
        {
            NDAWebClient client = null;
            try
            {
                client = new NDAWebClient();
                client.Timeout = timeout; 
                client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                byte[] data = client.DownloadData(pageUrl); 
                return data;
            }
            catch (Exception ex)
            {
                string errorInfo = "访问失败, PageUrl = " + pageUrl + ". " + ex.Message;
                //this.InvokeAppendLogText(errorInfo, LogLevelType.Error, true);
                throw new GrabRequestException(errorInfo);
            }
        }
        #endregion
    }
}
