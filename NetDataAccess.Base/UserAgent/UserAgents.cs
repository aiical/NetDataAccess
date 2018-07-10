using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetDataAccess.Base.UserAgent
{
    public class UserAgents
    {
        #region 构造函数
        public UserAgents()
        {
            this.Load();
        }
        #endregion

        #region UserAgentFilePath
        /// <summary>
        /// UserAgentFilePath
        /// </summary>
        private static string UserAgentFilePath = Path.Combine(Path.GetDirectoryName(Application.StartupPath), "Files/Config/UserAgent.xlsx");
        #endregion

        #region 所有UserAgent信息
        private List<string> _AllUserAgents = new List<string>();
        /// <summary>
        /// 所有UserAgent信息
        /// </summary>
        public List<string> AllUserAgents
        {
            get
            {
                return _AllUserAgents;
            }
        }

        private List<string> _AllPcUserAgents = new List<string>();
        /// <summary>
        /// 所有PC浏览器的UserAgent信息
        /// </summary>
        public List<string> AllPcUserAgents
        {
            get
            {
                return _AllPcUserAgents;
            }
        }

        private List<string> _AllMobileUserAgents = new List<string>();
        /// <summary>
        /// 所有Mobile浏览器的UserAgent信息
        /// </summary>
        public List<string> AllMobileUserAgents
        {
            get
            {
                return _AllMobileUserAgents;
            }
        }
        #endregion

        #region 获取一个UserAgent
        public string GetOneUserAgent()
        {
            int index = R.Next(AllUserAgents.Count);
            return AllUserAgents[index];

        }
        public string GetOnePcUserAgent()
        {
            int index = R.Next(AllPcUserAgents.Count);
            return AllUserAgents[index];

        }
        public string GetOneMobileUserAgent()
        {
            int index = R.Next(AllMobileUserAgents.Count);
            return AllUserAgents[index];

        }
        private Random R = new Random(DateTime.Now.Millisecond);
        #endregion

        #region 加载UserAgent列表
        /// <summary>
        /// 加载UserAgent列表
        /// </summary>
        /// <param name="filePath"></param>
        public void Load()
        {
            try
            {
                string filePath = UserAgentFilePath;
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook wb = new XSSFWorkbook(fs);
                    ISheet sheet = wb.GetSheet("UserAgent");
                    List<string> allUserAgents = new List<string>();
                    List<string> allPcUserAgents = new List<string>();
                    List<string> allMobileUserAgents = new List<string>();
                    for (int i = 1; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);

                        ICell uaCell = row.GetCell(0);
                        string ua = uaCell.ToString();
                        allUserAgents.Add(ua);

                        ICell clientCell = row.GetCell(1);
                        switch (clientCell.ToString())
                        {
                            case "pc":
                                allPcUserAgents.Add(ua);
                                break;
                            case "mobile":
                                allPcUserAgents.Add(ua);
                                break;
                        }
                    }
                    _AllUserAgents = allUserAgents;
                    _AllPcUserAgents = allPcUserAgents;
                    _AllMobileUserAgents = allMobileUserAgents;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("读取UserAgent出错", ex);
            }
        }
        #endregion

    }
}
