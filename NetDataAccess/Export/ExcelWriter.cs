using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.DB;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NetDataAccess.Export
{
    /// <summary>
    /// 保存到xlsx文档
    /// </summary>
    internal class ExcelWriter : DetailExportWriter
    {
        #region Workbook
        private IWorkbook _Workbook = null;
        #endregion

        #region Sheet
        private ISheet _DetailSheet = null;
        #endregion

        #region 输出文件路径
        private string _ExportFilePath = "";
        #endregion

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="exportFilePath">输出文件路径</param>
        /// <param name="columnNameToIndex">列名及列序号</param>
        public ExcelWriter(string exportFilePath, Dictionary<string, int> columnNameToIndex)
        {
            _ExportFilePath = exportFilePath;
            CommonUtil.CreateFileDirectory(exportFilePath);
            _Workbook = new XSSFWorkbook();
            _DetailSheet = this._Workbook.CreateSheet();
            IRow titileRow = this._DetailSheet.CreateRow(0);
            foreach (string columnName in columnNameToIndex.Keys)
            {
                int columnIndex = columnNameToIndex[columnName];
                ICell cell = titileRow.CreateCell(columnIndex);
                cell.SetCellValue(columnName);
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
            Dictionary<string,string> listRow = listSheet.GetRow(rowIndex);
            string urlCellValue = listRow[SysConfig.DetailPageUrlFieldName];
            if (urlCellValue == pageUrl)
            {
                IRow detailRow = this._DetailSheet.CreateRow(this._DetailSheet.LastRowNum + 1);
                foreach (string columnName in columnNameToIndex.Keys)
                {
                    int index = columnNameToIndex[columnName];
                    if (listRow.ContainsKey(columnName))
                    {
                        string v = listRow[columnName];
                        if (CommonUtil.IsNullOrBlank(v))
                        {
                            detailRow.CreateCell(index).SetCellValue(v);
                        }
                    }
                }

                foreach (string fieldName in fieldValues.Keys)
                {
                    int index = columnNameToIndex[fieldName];
                    string value = fieldValues[fieldName];
                    ICell cell = detailRow.CreateCell(index);
                    cell.SetCellValue(value);
                }
            }
            else
            {
                throw new Exception("第" + rowIndex.ToString() + "行地址不匹配. Url_1 = " + pageUrl + ", Url_2 = " + urlCellValue);
            }
        }
        #endregion         

        #region 保存到硬盘
        /// <summary>
        /// 保存到硬盘
        /// </summary>
        public void SaveToDisk()
        {
            FileStream fs = null;
            try
            {
                fs = new FileStream(this._ExportFilePath, FileMode.Create);
                this._Workbook.Write(fs);
                this._Workbook = null;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                    fs.Dispose();
                    fs = null;
                }
            }
        }
        #endregion    
    }
}
