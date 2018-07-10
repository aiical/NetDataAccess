using NetDataAccess.Base.Common;
using NetDataAccess.Base.CsvHelper;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace NetDataAccess.Base.Writer
{
    /// <summary>
    /// CsvWriter
    /// </summary>
    public class CsvWriter
    {
        #region 文件路径
        private string _FilePath = "";
        #endregion
        
        #region 临时保存在内存中的Datatable
        private DataTable _DataTable = null;
        #endregion

        #region 所有列
        private Dictionary<string, int> _ColumnNameToIndex = null;
        private Dictionary<string, int> ColumnNameToIndex
        {
            get
            {
                return this._ColumnNameToIndex;
            }
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="columnNameToIndex"></param>
        public CsvWriter(string filePath, Dictionary<string, int> columnNameToIndex)
        { 
            this._FilePath = filePath;
            this._DataTable = new DataTable();
            string[] columnNames = new string[columnNameToIndex.Count];
            foreach (string columnName in columnNameToIndex.Keys)
            {
                int index = columnNameToIndex[columnName];
                columnNames[index] = columnName; 
            }
            for (int i = 0; i < columnNames.Length; i++)
            {
                this._DataTable.Columns.Add(columnNames[i]);
            }
            this._ColumnNameToIndex = columnNameToIndex;
        }
        #endregion

        #region 设置行记录值
        public void SetFieldValues(int rowIndex, Dictionary<string, string> f2vs)
        {
            DataRow row = this._DataTable.Rows[rowIndex];
            foreach (string columnName in f2vs.Keys)
            { 
                row[columnName] = f2vs[columnName]; 
            }
        }
        #endregion

        #region 增加行记录
        public void AddRow(Dictionary<string, string> f2vs)
        {
            DataRow row = this._DataTable.NewRow();
            
            foreach (string columnName in f2vs.Keys)
            {
                row[columnName] = f2vs[columnName]; 
            }
            this._DataTable.Rows.Add(row);
        }
        #endregion

        #region 保存到硬盘
        public void SaveToDisk()
        {
            ElencyCsvWriter w = null;
            try
            {
                CommonUtil.CreateFileDirectory(this._FilePath);
                w = new ElencyCsvWriter();
                w.WriteCsv(this._DataTable, this._FilePath, Encoding.UTF8);
                
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (w != null)
                {
                    w.Dispose();
                }
            }
        }
        #endregion 
    }
}
