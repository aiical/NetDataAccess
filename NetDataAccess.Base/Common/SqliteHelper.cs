using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SQLite;
using System.Data.Common;
using System.IO;
using System.Windows.Forms; 

namespace NetDataAccess.Base.Common
{
    /// <summary>
    /// 本地数据库Sqlite的基本操作方法类
    /// </summary>
    public class SqliteHelper
    {
        #region 主数据库存储位置
        /// <summary>
        /// 主数据库存储位置
        /// </summary>
        private static string MainDbPath = Path.Combine(Path.GetDirectoryName(Application.StartupPath), "Files/Config/nda.db");
        #endregion

        #region 主数据库操作
        public static SqliteHelper _MainDbHelper = null;
        public static SqliteHelper MainDbHelper
        {
            get
            {
                if (_MainDbHelper == null)
                {
                    _MainDbHelper = new SqliteHelper(MainDbPath);
                }
                return _MainDbHelper;
            }
        }
        #endregion

        #region 数据库存储位置
        private string _SqliteDbPath = "";
        private string SqliteDbPath 
        {
            get
            {
                return _SqliteDbPath;
            } 
        }
        #endregion

        #region 构造函数
        public SqliteHelper(string sqliteDbPath)
        {
            this._SqliteDbPath = sqliteDbPath;
        }
        #endregion

        #region 获取数据库连接
        private SQLiteConnection connection = null;
        /// <summary>
        /// 获取数据库连接
        /// </summary>
        /// <returns></returns>
        private SQLiteConnection GetConnection()
        {
            if (connection == null)
            {
                string connStr = string.Format(@"Data Source={0}", SqliteDbPath);
                connection = new SQLiteConnection(connStr);
                connection.Open();
            }
            return connection;
        }
        #endregion

        #region 关闭
        public void Close()
        {
            if (this.connection != null && this.connection.State == ConnectionState.Open)
            {
                this.connection.Close();
                this.connection.Dispose();
            }
        }
        #endregion

        #region 根据sql和参数获取数据记录表
        public DataTable GetDataTable(string sql, Dictionary<string, object> p2vs)
        {
            SQLiteConnection conn = GetConnection();
            return this.GetDataTable(sql, p2vs, conn);
        }
        #endregion

        #region 根据sql和参数获取数据记录表
        public DataTable GetDataTable(string sql, Dictionary<string, object> p2vs, SQLiteConnection conn)
        {
            try
            { 
                SQLiteCommand cmd = new SQLiteCommand(conn);
                cmd.CommandText = sql;
                if (p2vs != null)
                {
                    foreach (string pName in p2vs.Keys)
                    {
                        cmd.Parameters.AddWithValue(pName, p2vs[pName]);
                    }
                }
                SQLiteDataAdapter dao = new SQLiteDataAdapter(cmd);
                DataTable dt = new DataTable();
                dao.Fill(dt);
                return dt;
            }
            catch (Exception ex)
            {
                LogHelper.WriteMessage(sql + "\r\n" + ex.Message);
                throw ex;
            }
        }
        #endregion

        #region 执行insert,update,delete动作
        public bool ExecuteSql(string sql)
        {
            SQLiteConnection conn = GetConnection();
            return this.ExecuteSql(sql, conn);
        }
        #endregion

        #region 执行insert,update,delete动作
        public bool ExecuteSql(string sql, SQLiteConnection conn)
        {
            try
            {
                SQLiteCommand cmd = new SQLiteCommand(conn);
                cmd.CommandText = sql; 
                cmd.ExecuteNonQuery();
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteMessage(sql + "\r\n" + ex.Message);
                throw ex;
            }
        }
        #endregion

        #region 执行insert,update,delete动作
        /// <summary>
        /// 执行insert,update,delete动作
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="p2vs"></param>
        /// <returns></returns>
        public bool ExecuteSql(string sql, Dictionary<string, object> p2vs)
        {
            SQLiteConnection conn = GetConnection();
            return this.ExecuteSql(sql, p2vs, conn);
        }
        #endregion

        #region 执行insert,update,delete动作
        /// <summary>
        /// 执行insert,update,delete动作
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="p2vs"></param>
        /// <returns></returns>
        public bool ExecuteSql(string sql, Dictionary<string, object> p2vs, SQLiteConnection conn)
        {
            try
            {
                SQLiteCommand cmd = new SQLiteCommand(conn);
                cmd.CommandText = sql;
                if (p2vs != null)
                {
                    foreach (string pName in p2vs.Keys)
                    {
                        cmd.Parameters.AddWithValue(pName, p2vs[pName]);
                    }
                }
                cmd.ExecuteNonQuery();
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteMessage(sql + "\r\n" + ex.Message);
                throw ex;
            }
        }
        #endregion

        #region 执行insert,update,delete动作
        /// <summary>
        /// 执行insert,update,delete动作
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="p2vs"></param>
        /// <returns></returns>
        public bool ExecuteSql(string sql, Dictionary<string, string> p2vs)
        {
            SQLiteConnection conn = GetConnection();
            return this.ExecuteSql(sql, p2vs, conn);
        }
        #endregion

        #region 执行insert,update,delete动作
        /// <summary>
        /// 执行insert,update,delete动作
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="p2vs"></param>
        /// <returns></returns>
        public bool ExecuteSql(string sql, Dictionary<string, string> p2vs, SQLiteConnection conn)
        {
            try
            { 
                SQLiteCommand cmd = new SQLiteCommand(conn);
                cmd.CommandText = sql;
                if (p2vs != null)
                {
                    foreach (string pName in p2vs.Keys)
                    {
                        cmd.Parameters.AddWithValue(pName, p2vs[pName]);
                    }
                }
                cmd.ExecuteNonQuery();
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteMessage(sql + "\r\n" + ex.Message);
                throw ex;
            }
        }
        #endregion

        #region 创建Command
        public SQLiteCommand CreateCommand(SQLiteConnection conn)
        {
            SQLiteCommand cmd = new SQLiteCommand((SQLiteConnection)conn);
            return cmd;
        }
        #endregion

        #region 执行insert,update,delete动作
        /// <summary>
        /// 执行insert,update,delete动作
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="allP2vs"></param>
        /// <returns></returns>
        public bool ExecuteSql(string sql, List<Dictionary<string, string>> allP2vs)
        {
            SQLiteConnection conn = GetConnection();
            return this.ExecuteSql(sql, allP2vs, conn);
        }
        #endregion

        #region 执行insert,update,delete动作
        /// <summary>
        /// 执行insert,update,delete动作
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="allP2vs"></param>
        /// <returns></returns>
        public bool ExecuteSql(string sql, List<Dictionary<string, string>> allP2vs, SQLiteConnection conn)
        {
            SQLiteTransaction tran = null;
            try
            { 
                for (int i = 0; i < allP2vs.Count; i++)
                {
                    bool needCommit = i % 100000 == 0;
                    if (needCommit)
                    {
                        if (i != 0)
                        {
                            tran.Commit();
                        }
                        tran = conn.BeginTransaction();
                    }
                    SQLiteCommand cmd = new SQLiteCommand(conn);
                    cmd.CommandText = sql;
                    Dictionary<string, string> p2vs = allP2vs[i];
                    if (p2vs != null)
                    {
                        foreach (string pName in p2vs.Keys)
                        {
                            cmd.Parameters.AddWithValue(pName, p2vs[pName]);
                        }
                    }
                    cmd.ExecuteNonQuery();
                }
                tran.Commit();
                return true;
            }
            catch (Exception ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                LogHelper.WriteMessage(sql + "\r\n" + ex.Message);
                throw ex;
            }
        }
        #endregion

        #region 执行事务
        public delegate object RunTransactionDelegate(SQLiteConnection dbConnection);
        public object RunTransaction(RunTransactionDelegate run)
        {
            SQLiteTransaction tran = null;
            try
            {
                SQLiteConnection conn = GetConnection();
                tran = conn.BeginTransaction();
                Object obj = run(conn);
                tran.Commit();
                return obj;
            }
            catch (Exception ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                LogHelper.WriteMessage(ex.Message);
                throw ex;
            }
        }
        #endregion

        #region 执行insert,update,delete动作
        /// <summary>
        /// 执行insert,update,delete动作
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="allP2vs"></param>
        /// <returns></returns>
        public bool ExecuteSql(string sql, List<Dictionary<string, object>> allP2vs)
        {
            SQLiteConnection conn = GetConnection();
            return this.ExecuteSql(sql, allP2vs, conn);
        }
        #endregion

        #region 执行insert,update,delete动作
        /// <summary>
        /// 执行insert,update,delete动作
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="allP2vs"></param>
        /// <returns></returns>
        public bool ExecuteSql(string sql, List<Dictionary<string, object>> allP2vs, SQLiteConnection conn)
        {
            SQLiteTransaction tran = null;
            try
            { 
                for (int i = 0; i < allP2vs.Count; i++)
                {
                    bool needCommit = i % 1000 == 0;
                    if (needCommit)
                    {
                        if (i != 0)
                        {
                            tran.Commit();
                        }
                        tran = conn.BeginTransaction();
                    }
                    SQLiteCommand cmd = new SQLiteCommand(conn);
                    cmd.CommandText = sql;
                    Dictionary<string, object> p2vs = allP2vs[i];
                    if (p2vs != null)
                    {
                        foreach (string pName in p2vs.Keys)
                        {
                            cmd.Parameters.AddWithValue(pName, p2vs[pName]);
                        }
                    }
                    cmd.ExecuteNonQuery();
                }
                tran.Commit();
                return true;
            }
            catch (Exception ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                LogHelper.WriteMessage(sql + "\r\n" + ex.Message);
                throw ex;
            }
        }
        #endregion

        #region 执行select，包含单个返回值
        /// <summary>
        /// 执行select，包含单个返回值
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="p2vs"></param>
        /// <returns></returns>
        public object ExecuteSelectOneSql(string sql, Dictionary<string, string> p2vs)
        {
            SQLiteConnection conn = GetConnection();
            return this.ExecuteSelectOneSql(sql, p2vs, conn);
        }
        #endregion

        #region 执行select，包含单个返回值
        /// <summary>
        /// 执行select，包含单个返回值
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="p2vs"></param>
        /// <returns></returns>
        public object ExecuteSelectOneSql(string sql, Dictionary<string, string> p2vs, SQLiteConnection conn)
        {
            try
            {
                SQLiteCommand cmd = new SQLiteCommand(conn);
                cmd.CommandText = sql;
                if (p2vs != null)
                {
                    foreach (string pName in p2vs.Keys)
                    {
                        cmd.Parameters.AddWithValue(pName, p2vs[pName]);
                    }
                }
                return cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                LogHelper.WriteMessage(sql + "\r\n" + ex.Message);
                throw ex;
            }
        }
        #endregion
    }
}