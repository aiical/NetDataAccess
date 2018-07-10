using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.Web;
using NetDataAccess.Base.Writer;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace NetDataAccess.Base.Proxy
{
    /// <summary>
    /// 代理服务器集合
    /// </summary>
    public class ProxyServers
    {
        #region 构造函数
        public ProxyServers()
        {
            this.Load();
        }
        #endregion

        #region ServerListFilePath
        /// <summary>
        /// ServerListFilePath
        /// </summary>
        private static string ServerListFilePath = Path.Combine(Path.GetDirectoryName(Application.StartupPath), "Files/Config/Proxy.xlsx");
        #endregion

        #region 锁
        private object _GetProxyServerLocker = new object();
        #endregion

        #region 加载代理服务器列表
        /// <summary>
        /// 加载代理服务器列表
        /// </summary>
        /// <param name="filePath"></param>
        public void Load()
        {
            try
            {
                string filePath = ServerListFilePath;
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook wb = new XSSFWorkbook(fs);
                    ISheet sheet = wb.GetSheet("Proxy");
                    List<ProxyServer> allProxies = new List<ProxyServer>();
                    for (int i = 1; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);

                        ICell usableCell = row.GetCell(4);
                        if (usableCell != null && usableCell.StringCellValue == "是")
                        {
                            ProxyServer ps = new ProxyServer();
                            ps.Index = i - 1;
                            ps.IP = row.GetCell(0).ToString();
                            ps.Port = int.Parse(row.GetCell(1).ToString());
                            ps.User = row.GetCell(2) == null ? "" : row.GetCell(2).ToString();
                            ps.Pwd = row.GetCell(3) == null ? "" : row.GetCell(3).ToString();
                            allProxies.Add(ps);
                        }
                    }
                    _AllProxies = allProxies;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("读取代理服务器出错", ex);
            }
        }
        #endregion

        #region 所有代理服务器信息
        private List<ProxyServer> _AllProxies = new List<ProxyServer>();
        /// <summary>
        /// 所有代理服务器信息
        /// </summary>
        public List<ProxyServer> AllProxies
        {
            get {
                return _AllProxies;
            }
        }
        #endregion

        #region 获取一个可用的代理服务器

        public int _CurrentCheckIndex = 0;

        /// <summary>
        /// 获取一个可用的代理服务器
        /// </summary>
        /// <param name="needIdleTime">需要使用时间，毫秒</param>
        /// <returns></returns>
        public ProxyServer BeginUse(int useTime)
        {
            lock (_GetProxyServerLocker)
            {
                int totalCount = this.AllProxies.Count; 
                for (int i = 0; i < totalCount; i++)
                {
                    _CurrentCheckIndex++;
                    _CurrentCheckIndex = _CurrentCheckIndex < totalCount ? _CurrentCheckIndex : (_CurrentCheckIndex - totalCount);
                    ProxyServer ps = this.AllProxies[_CurrentCheckIndex];
                    if (!ps.IsAbandon && ps.AvailableTime < DateTime.Now)
                    { 
                        ps.AvailableTime = DateTime.Now.AddMilliseconds(useTime);
                        return ps;
                    }
                }
                throw new NoneProxyException("暂无可用的代理服务器.");
            }
        }
        #endregion

        #region 获取一个可用的代理服务器
        public void EndUse(ProxyServer ps)
        { 
        }
        #endregion

        #region 获取一个可用的代理服务器
        public int GetAvailableCount()
        {
            lock (_GetProxyServerLocker)
            {
                int count = 0;
                foreach (ProxyServer ps in this.AllProxies)
                {
                    if ( !ps.IsAbandon)
                    {
                        count++;
                    }
                }
                return count;
            }
        }
        #endregion

        #region 记录正常使用过的代理
        /* 作废
        private static List<int> UsedProxyServers = new List<int>();

        public static void AddUsedProxyServer(int index)
        {
            if (!UsedProxyServers.Contains(index))
            {
                UsedProxyServers.Add(index);
            }
        }

        public static void SaveUsedProxyServers()
        { 
            string filePath = Path.Combine(Path.GetDirectoryName(Application.StartupPath), "Files/Config/UsedProxy" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
            Dictionary<string, int> columns = new Dictionary<string, int>();
            columns.Add("ip", 0);
            columns.Add("port", 1);
            columns.Add("user", 2);
            columns.Add("pwd", 3);
            columns.Add("usable", 4);
            ExcelWriter ew = new ExcelWriter(filePath, "Proxy", columns);

            StringBuilder sb = new StringBuilder();
            foreach (int index in UsedProxyServers)
            {
                ProxyServer ps = AllProxies[index];
                Dictionary<string, string> p2vs = new Dictionary<string, string>();
                p2vs.Add("ip", ps.IP);
                p2vs.Add("port", ps.Port.ToString());
                p2vs.Add("user", ps.User);
                p2vs.Add("pwd", ps.Pwd);
                p2vs.Add("usable", "是");
                ew.AddRow(p2vs);
            }
            ew.SaveToDisk();
            CommonUtil.Alert("保存成功", "包含" + UsedProxyServers.Count.ToString() + "个可用代理服务器，保存地址为" + filePath);
        }
        */
        #endregion

        #region 成功
        public void Success(ProxyServer ps)
        {
            if (ps != null)
            {
                lock (_GetProxyServerLocker)
                {
                    ps.ErrorCount = 0; 
                }
            }
        }
        #endregion

        #region 放弃
        public void Error(ProxyServer ps)
        {
            if (ps != null)
            {
                lock (_GetProxyServerLocker)
                {
                    ps.AddErrorCount();
                    if (SysConfig.ProxyAbandonErrorTime != 0 && ps.ErrorCount >= SysConfig.ProxyAbandonErrorTime)
                    {
                        ps.IsAbandon = true;
                    }
                }
            }
        }
        #endregion   
    }
}
