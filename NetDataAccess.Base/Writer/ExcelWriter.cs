using NetDataAccess.Base.Common;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NetDataAccess.Base.Writer
{
    /// <summary>
    /// ExcelWriter
    /// </summary>
    public class ExcelWriter
    {
        #region 文件路径
        private string _FilePath = "";
        #endregion

        #region Workbook
        private IWorkbook _Workbook = null;
        #endregion

        #region Sheet
        private ISheet _Sheet = null;
        public ISheet Sheet 
        {
            get
            {
                return this._Sheet;
            }
        }
        #endregion

        #region 所有列样式
        private Dictionary<string, ICellStyle> _ColumnNameToCellStyle = null;
        private Dictionary<string, ICellStyle> ColumnNameToCellStyle
        {
            get
            {
                return this._ColumnNameToCellStyle;
            }
        }
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

        #region 获取行
        public IRow GetRow(int rowIndex)
        {
            IRow row = this._Sheet.GetRow(rowIndex + 1);
            return row;
        }
        #endregion

        #region 行数 
        private int _RowCount = 0;
        public int RowCount
        {
            get
            {
                return _RowCount;
            }
        } 
        #endregion

        #region 获取单元格
        public ICell GetCell(int rowIndex, string columnName, bool autoCreate)
        {
            IRow row = this._Sheet.GetRow(rowIndex + 1);
            ICell cell = GetCell(row, columnName, autoCreate);
            return cell;
        }
        #endregion

        #region 获取单元格
        public ICell GetCell(IRow row, string columnName, bool autoCreate)
        { 
            int cellIndex = this._ColumnNameToIndex[columnName];
            ICell cell = row.GetCell(cellIndex);
            if (cell == null && autoCreate)
            {
                cell = row.CreateCell(cellIndex);
            }
            return cell;
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        /// <param name="columnNameToIndex"></param>
        public ExcelWriter(string filePath, string sheetName, Dictionary<string, int> columnNameToIndex)
        {
            this.Create(filePath, sheetName, columnNameToIndex, null, null);
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        /// <param name="columnNameToIndex"></param>
        /// <param name="columnNameToFormat"></param>
        public ExcelWriter(string filePath, string sheetName, Dictionary<string, int> columnNameToIndex, Dictionary<string, string> columnNameToFormat)
        {
            this.Create(filePath, sheetName, columnNameToIndex, columnNameToFormat, null);
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        /// <param name="columnNameToIndex"></param>
        /// <param name="columnNameToFormat"></param>
        public ExcelWriter(string filePath, string sheetName, List<object[]> columns)
        {
            this.Create(filePath, sheetName, columns);
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        /// <param name="columnNameToIndex"></param>
        /// <param name="columnNameToFormat"></param>
        public ExcelWriter(string filePath, string sheetName, Dictionary<string, int> columnNameToIndex, Dictionary<string, string> columnNameToFormat, Dictionary<String, int> columnNameToWidth)
        {
            this.Create(filePath, sheetName, columnNameToIndex, columnNameToFormat, columnNameToWidth);
        }
        #endregion

        #region 创建
        private void Create(string filePath, string sheetName, List<object[]> columns)
        {
            this._FilePath = filePath;
            this._Workbook = new XSSFWorkbook();
            this._Sheet = this._Workbook.CreateSheet(sheetName);
            this._ColumnNameToCellStyle = new Dictionary<string, ICellStyle>();

            IRow titleRow = this._Sheet.CreateRow(0);
            Dictionary<string, int> columnNameToIndex = new Dictionary<string, int>();
            for (int index = 0; index < columns.Count; index++)
            {
                object[] column = columns[index];
                string columnName = (string)column[0];
                ICell cell = titleRow.CreateCell(index);
                cell.SetCellValue(columnName);
                columnNameToIndex.Add(columnName, index);

                string format = (string)column[1];
                if (format != null && format.Length != 0)
                {
                    ICellStyle cellStyle = null;
                    IDataFormat dataFormat = null;
                    cellStyle = this._Workbook.CreateCellStyle();
                    dataFormat = this._Workbook.CreateDataFormat();
                    cellStyle.DataFormat = dataFormat.GetFormat(format);
                    this._ColumnNameToCellStyle.Add(columnName, cellStyle);
                }

                int columnWidth = (int)column[2];
                this.Sheet.SetColumnWidth(index, columnWidth * 256);
            }
            this._ColumnNameToIndex = columnNameToIndex;

        }
        #endregion

        #region 创建
        private void Create(string filePath, string sheetName, Dictionary<string, int> columnNameToIndex, Dictionary<string, string> columnNameToFormat, Dictionary<string, int> columnNameToWidth)
        {
            this._FilePath = filePath;
            this._Workbook = new XSSFWorkbook();
            this._Sheet = this._Workbook.CreateSheet(sheetName);

            IRow titleRow = this._Sheet.CreateRow(0);
            foreach (string columnName in columnNameToIndex.Keys)
            {
                int index = columnNameToIndex[columnName];
                ICell cell = titleRow.CreateCell(index);
                cell.SetCellValue(columnName);
            }
            this._ColumnNameToIndex = columnNameToIndex;

            if (columnNameToWidth != null)
            {
                foreach (string columnName in columnNameToWidth.Keys)
                {
                    int columnIndex = this.ColumnNameToIndex[columnName];
                    int columnWidth = columnNameToWidth[columnName];
                    this.Sheet.SetColumnWidth(columnIndex, columnWidth * 256);
                }
            }
            
            this._ColumnNameToCellStyle = new Dictionary<string, ICellStyle>();
            if (columnNameToFormat != null)
            {
                foreach (string columnName in columnNameToFormat.Keys)
                {
                    string format = columnNameToFormat[columnName];
                    ICellStyle cellStyle = null;
                    IDataFormat dataFormat = null;
                    if (format != null)
                    {
                        cellStyle = this._Workbook.CreateCellStyle();
                        dataFormat = this._Workbook.CreateDataFormat();
                        cellStyle.DataFormat = dataFormat.GetFormat(format);
                        this._ColumnNameToCellStyle.Add(columnName, cellStyle);
                    }
                }
            }
        }
        #endregion

        #region 设置行记录值
        public void SetFieldValues(int rowIndex, Dictionary<string, object> f2vs)
        {
            IRow row = this._Sheet.GetRow(rowIndex + 1);
            foreach (string columnName in f2vs.Keys)
            {
                int index = this._ColumnNameToIndex[columnName];
                ICell cell = row.GetCell(index);
                if (cell == null)
                {
                    cell = row.CreateCell(index);
                }
                ICellStyle cellStyle = this._ColumnNameToCellStyle.ContainsKey(columnName) ? this._ColumnNameToCellStyle[columnName] : null;
                object valueObj = f2vs[columnName];

                cell.SetCellType(CellType.Formula);
                cell.CellStyle = cellStyle;
                if (valueObj != null)
                {
                    if (valueObj is DateTime)
                    {
                        cell.SetCellValue((DateTime)valueObj);
                    }
                    else if (valueObj is decimal)
                    {
                        cell.SetCellValue((double)valueObj);
                    }
                }
            }
        }
        #endregion

        #region 增加行记录
        public IRow AddRow(Dictionary<string, string> f2vs)
        {
            _RowCount++;
            IRow row = this._Sheet.CreateRow(this._Sheet.LastRowNum + 1);
            foreach (string columnName in f2vs.Keys)
            {
                int index = this._ColumnNameToIndex[columnName];
                ICell cell = row.CreateCell(index);
                cell.SetCellValue(f2vs[columnName]);
            }
            return row;
        }
        #endregion

        #region 增加行记录
        public IRow AddRow(Dictionary<string, object> f2vs)
        {
            _RowCount++;
            IRow row = this._Sheet.CreateRow(this._Sheet.LastRowNum + 1);
            foreach (string columnName in f2vs.Keys)
            {
                int index = this._ColumnNameToIndex[columnName];
                ICell cell = row.CreateCell(index);
                ICellStyle cellStyle = this._ColumnNameToCellStyle.ContainsKey(columnName) ? this._ColumnNameToCellStyle[columnName] : null;
                object valueObj = f2vs[columnName];
                if (cellStyle == null)
                {
                    cell.SetCellValue(valueObj == null ? "" : valueObj.ToString());
                }
                else
                {
                    cell.CellStyle = cellStyle;
                    if (valueObj != null)
                    {
                        if (valueObj is DateTime)
                        {
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue((DateTime)valueObj);
                        }
                        else if (valueObj is decimal || valueObj is double || valueObj is int)
                        {
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(double.Parse(valueObj.ToString()));
                        }
                        else
                        {
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(double.Parse(valueObj.ToString()));
                        }
                    }
                }
            }
            return row;
        }
        #endregion

        #region 关闭文件
        public void Close()
        {
            this._Workbook.Close();
        }
        #endregion

        #region 保存到硬盘
        public void SaveToDisk()
        {
            FileStream fs = null;
            try
            {
                CommonUtil.CreateFileDirectory(this._FilePath);
                fs = new FileStream(this._FilePath, FileMode.Create);
                this._Workbook.Write(fs);
                this._Workbook.Close();
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
                }
            }
        }
        #endregion
    }
}
