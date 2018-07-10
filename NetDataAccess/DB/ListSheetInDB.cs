using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.DB; 
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace NetDataAccess.DB
{
    /// <summary>
    /// Sqlite中保存读取要获取的网址信息
    /// </summary>
    public class ListSheetInDB : IListSheet
    {
        #region 锁
        private object _DBLocker = new object();
        #endregion

        #region 行数
        private int _RowCount = 0;
        public int RowCount
        {
            get
            {
                return this._RowCount;
            }
        }
        #endregion

        #region 本地数据库Sqlite的基本操作方法类
        private SqliteHelper _DBHelper = null;
        private SqliteHelper DBHelper
        {
            get
            {
                return this._DBHelper;
            }
        }
        #endregion

        #region 网址列表
        private List<string> _PageUrlList = null;
        public List<string> PageUrlList
        {
            get
            {
                return _PageUrlList;
            }
        }
        #endregion

        #region Cookie列表
        private List<string> _PageCookieList = null;
        public List<string> PageCookieList
        {
            get
            {
                return _PageCookieList;
            }
        }
        #endregion

        #region 网址名称列表
        private List<string> _PageNameList = null;
        public List<string> PageNameList
        {
            get
            {
                return _PageNameList;
            }
        }
        #endregion

        #region 记录被放弃读取（按顺序记录，true为放弃，false为需要爬取）
        private List<bool> _GiveUpList = null;
        public List<bool> GiveUpList
        {
            get
            {
                return _GiveUpList;
            }
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="dbFilePath"></param>
        public ListSheetInDB(string dbFilePath)
        {
            _DBHelper = new SqliteHelper(dbFilePath);
        }
        #endregion

        #region 关闭数据库连接
        public void Close()
        {
            this.DBHelper.Close();
        }
        #endregion

        #region 从数据库中加载详情页列表信息，并初始化到内存
        public void InitDetailPageInfo()
        {
            string sql = "select " + SysConfig.DetailPageUrlFieldName + ", " + SysConfig.DetailPageNameFieldName + ", " + SysConfig.DetailPageCookieFieldName + ", " + SysConfig.GiveUpGrabFieldName + " from list";// where " + SysConfig.GiveUpGrabFieldName + " = 'N'";
            DataTable dt = this.DBHelper.GetDataTable(sql, null);
            this._PageNameList = new List<string>();
            this._PageUrlList = new List<string>();
            this._PageCookieList = new List<string>();
            this._GiveUpList = new List<bool>();
            this._RowCount = dt.Rows.Count;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow row = dt.Rows[i];
                string url = (string)row[SysConfig.DetailPageUrlFieldName];
                string name = (string)row[SysConfig.DetailPageNameFieldName];
                string cookie = (string)row[SysConfig.DetailPageCookieFieldName];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                this.PageNameList.Add(name);
                this.PageUrlList.Add(url);
                this.PageCookieList.Add(cookie);
                this.GiveUpList.Add(giveUp);
            }
        }
        #endregion

        #region 获取已经抓取到的列表页序号，列表页是顺序爬取的，有一个页面出错后就立刻停止，所以获取到的count(1)就是列表页序号
        public int GetSucceedListPageIndex()
        {
            string sql = "select count(1) from list";
            object count = this.DBHelper.ExecuteSelectOneSql(sql, null);
            return int.Parse(count.ToString());
        }
        #endregion

        #region 获取ListPageIndex的最大值
        public int GetListDBRowCount()
        {
            string sql = "select count(" + SysConfig.ListPageIndexFieldName + ") from list";
            object count = this.DBHelper.ExecuteSelectOneSql(sql, null);
            return int.Parse(count.ToString());
        }
        #endregion

        #region 增加列
        public void AddColumns(List<string> newColumns)
        {
            foreach (string column in newColumns)
            {
                string sql = "alter table list add " + column + " varchar(512);";
                this.DBHelper.ExecuteSql(sql);
            }
        }
        #endregion

        #region 将要爬取的网址插入到数据库
        public void CopyToListSheet(List<Dictionary<string, string>> allF2vs, Dictionary<string, int> columnToIndexs, List<string> sysColumnList)
        {
            StringBuilder addSqlField = new StringBuilder(SysConfig.ListPageIndexFieldName + ","
                               + SysConfig.DetailPageUrlFieldName + ","
                               + SysConfig.DetailPageNameFieldName + ","
                               + SysConfig.DetailPageCookieFieldName + ","
                               + SysConfig.GrabStatusFieldName + ","
                               + SysConfig.GiveUpGrabFieldName);

            StringBuilder addSqlValues = new StringBuilder(":" + SysConfig.ListPageIndexFieldName + ","
                               + ":" + SysConfig.DetailPageUrlFieldName + ","
                               + ":" + SysConfig.DetailPageNameFieldName + ","
                               + ":" + SysConfig.DetailPageCookieFieldName + ","
                               + ":" + SysConfig.GrabStatusFieldName + ","
                               + ":" + SysConfig.GiveUpGrabFieldName + " ");

            foreach (string columnName in columnToIndexs.Keys)
            {
                if (!sysColumnList.Contains(columnName))
                {
                    addSqlField.Append(", " + columnName);
                    addSqlValues.Append(", :" + columnName);
                }
            }
            string addSql = "insert into list(" + addSqlField.ToString() + ") values(" + addSqlValues + ")";
            
            this.DBHelper.ExecuteSql(addSql, allF2vs);
        }
        #endregion

        #region 设置放弃爬取
        public void SetGiveUp(int pageIndex, string pageUrl, string errorMsg)
        {
            string url = this.PageUrlList[pageIndex];
            if (url == pageUrl)
            {
                string sql = "update list set giveUpGrab = 'Y' where " + SysConfig.ListPageIndexFieldName + " = :" + SysConfig.ListPageIndexFieldName;
                Dictionary<string, object> p2vs = new Dictionary<string, object>();
                p2vs.Add(SysConfig.ListPageIndexFieldName, pageIndex.ToString());
                this.DBHelper.ExecuteSql(sql, p2vs);
                this.GiveUpList[pageIndex] = true;
            }
            else
            {
                throw new Exception("记录行定位错误. PageUrl = " + pageUrl + ". UrlCellValue = " + url);
            }
        }
        #endregion

        #region 增加一行
        public void AddListRow(Dictionary<string, string> f2vs)
        {
            StringBuilder addSqlField = new StringBuilder();

            StringBuilder addSqlValues = new StringBuilder();

            int fIndex = 0;
            foreach (string columnName in f2vs.Keys)
            {
                if (fIndex == 0)
                {
                    addSqlField.Append(columnName);
                    addSqlValues.Append(":" + columnName);
                }
                else
                {
                    addSqlField.Append(", " + columnName);
                    addSqlValues.Append(", :" + columnName);
                }
                fIndex++;
            }
            string addSql = "insert into list(" + addSqlField.ToString() + ") values(" + addSqlValues + ")";

            lock (_DBLocker)
            {
                f2vs[SysConfig.GiveUpGrabFieldName] = "N";
                this._RowCount++;
                this.PageNameList.Add(f2vs[SysConfig.DetailPageNameFieldName]);
                this.PageUrlList.Add(f2vs[SysConfig.DetailPageUrlFieldName]);
                this.GiveUpList.Add(false);
                this.DBHelper.ExecuteSql(addSql, f2vs);
            }
        }
        #endregion

        #region 获取一行
        public Dictionary<string, string> GetRow(int rowIndex)
        {
            string sql = "select * from list where " + SysConfig.ListPageIndexFieldName + " = :" + SysConfig.ListPageIndexFieldName;
            Dictionary<string, object> p2vs = new Dictionary<string, object>();
            p2vs.Add(SysConfig.ListPageIndexFieldName, rowIndex.ToString());
            DataTable dt = this.DBHelper.GetDataTable(sql, p2vs);
            Dictionary<string, string> f2vs =null;
            if (dt.Rows.Count > 0)
            {
                f2vs = new Dictionary<string, string>();
                DataRow row = dt.Rows[0];
                foreach (DataColumn c in dt.Columns)
                {
                    f2vs.Add(c.ColumnName, (string)row[c]);
                }
            }
            return f2vs;
        }
        #endregion
    }
}
