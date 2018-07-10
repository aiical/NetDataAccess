using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.Reader
{
    /// <summary>
    /// ExcelReader
    /// </summary>
    public class ExcelReader : TableFileReader, ITableFileReader
    {
        #region Workbook
        private IWorkbook _Workbook = null;
        #endregion

        #region Sheet
        private ISheet _Sheet = null;
        #endregion

        #region 所有列
        private Dictionary<string, int> _ColumnNameToIndex = null;
        public Dictionary<string, int> ColumnNameToIndex
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
        public ExcelReader(string filePath)
        {
            this._Workbook = new XSSFWorkbook(filePath);
            this._Sheet = this._Workbook.GetSheetAt(0);
            this.GetColumns();
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        public ExcelReader(string filePath, string sheetName)
        {
            this._Workbook = new XSSFWorkbook(filePath);
            this._Sheet = this._Workbook.GetSheet(sheetName);
            this.GetColumns();
        }
        #endregion

        #region 关闭
        public void Close()
        {
            this._Workbook.Close(); 
            this._Sheet = null; 
            this._Workbook = null;
            GC.Collect();
        }
        #endregion

        #region 根据首行信息获取列
        private void GetColumns()
        {
            IRow titleRow = this._Sheet.GetRow(0);
            int lastNum = titleRow.LastCellNum;
            Dictionary<string, int> columnNameToIndex = new Dictionary<string, int>();
            for (int i = 0; i <= lastNum; i++)
            {
                ICell cell = titleRow.GetCell(i);
                if (cell != null)
                {
                    columnNameToIndex.Add(cell.ToString(), i);
                }
            }
            this._ColumnNameToIndex = columnNameToIndex;
        }
        #endregion

        #region GetRowCount
        public override int GetRowCount()
        {
            return this._Sheet.LastRowNum;
        }
        #endregion

        #region GetFieldValues
        public override Dictionary<string, string> GetFieldValues(int rowIndex)
        {
            IRow row = this._Sheet.GetRow(rowIndex + 1);
            if (row != null)
            {
                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                foreach (string columnName in this._ColumnNameToIndex.Keys)
                {
                    int index = this._ColumnNameToIndex[columnName];
                    ICell cell = row.GetCell(index);
                    string value = cell == null ? "" : cell.ToString();
                    f2vs.Add(columnName, value);
                }
                return f2vs;
            }
            else
            {
                return null;
            }
        }
        #endregion
    }
}
