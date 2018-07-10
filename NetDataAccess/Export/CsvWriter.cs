using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.CsvHelper;
using NetDataAccess.Base.DB;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace NetDataAccess.Export
{
    /// <summary>
    /// 写入CSV文档
    /// </summary>
    internal class CsvWriter : DetailExportWriter
    {
        #region 输出文件路径
        private string _ExportFilePath = "";
        #endregion

        #region 暂存到内存
        private DataTable _DataTable = null;
        #endregion

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="exportFilePath">输出文件路径</param>
        /// <param name="columnNameToIndex">列名及列序号</param>
        public CsvWriter(string exportFilePath, Dictionary<string, int> columnNameToIndex)
        {
            this._ExportFilePath = exportFilePath;
            this._DataTable = new DataTable();
            
            //表头
            string[] allValues = new string[columnNameToIndex.Count];
            foreach (string fieldName in columnNameToIndex.Keys)
            {
                int index = columnNameToIndex[fieldName];
                allValues[index] = fieldName;
            }

            for (int i = 0; i < allValues.Length; i++)
            {
                this._DataTable.Columns.Add(allValues[i]);
            }
        }
        #endregion

        #region 保存一条记录
        public void AddRows(List<Dictionary<string, string>> valuesList)
        {
            foreach(Dictionary<string,string> values in valuesList)
            {
                DataRow row = this._DataTable.NewRow();
                foreach (string f in values.Keys)
                {
                    row[f] = values[f];
                }
                this._DataTable.Rows.Add(row);
            }
        }
        #endregion

        #region 保存一条记录
        public void AddRows(string fieldName, List<string> values)
        {
            foreach (string value in values)
            {
                DataRow row = this._DataTable.NewRow();
                row[fieldName] = value;
                this._DataTable.Rows.Add(row);
            }
        }
        #endregion

        #region 保存一条记录
        /// <summary>
        /// 保存一条记录
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="columnNameToIndex"></param>
        /// <param name="fieldValues"></param>
        /// <param name="rowIndex"></param>
        /// <param name="pageUrl"></param>
        public void SaveDetailFieldValue(IListSheet listSheet, Dictionary<string, int> columnNameToIndex, Dictionary<string, string> fieldValues, int rowIndex, string pageUrl)
        {
            try
            {
                Dictionary<string,string> listRow = listSheet.GetRow(rowIndex );
                string urlCellValue = listRow[SysConfig.DetailPageUrlFieldName];
                if (urlCellValue == pageUrl)
                { 
                    foreach (string columnName in columnNameToIndex.Keys)
                    {
                        if (listRow.ContainsKey(columnName))
                        {
                            string v = listRow[columnName];
                            if (!fieldValues.ContainsKey(columnName))
                            {
                                fieldValues.Add(columnName, v);
                            }
                        }
                    }

                    DataRow row = this._DataTable.NewRow();
                    foreach(string f in fieldValues.Keys)
                    {
                        row[f] = fieldValues[f];
                    }
                    this._DataTable.Rows.Add(row);
                }
                else
                {
                    throw new Exception("第" + rowIndex.ToString() + "行地址不匹配. Url_1 = " + pageUrl + ", Url_2 = " + urlCellValue);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            } 
        }
        #endregion

        #region 保存到硬盘
        /// <summary>
        /// 保存到硬盘
        /// </summary>
        public void SaveToDisk()
        {
            ElencyCsvWriter ecw = null;
            try
            {
                CommonUtil.CreateFileDirectory(this._ExportFilePath);
                ecw = new ElencyCsvWriter();
                ecw.WriteCsv(this._DataTable, this._ExportFilePath, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (ecw != null)
                {
                    ecw.Dispose();
                }
            }
        }
        #endregion
    }
}
