using NetDataAccess.Base.CsvHelper;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace NetDataAccess.Base.Reader
{
    /// <summary>
    /// CsvReader
    /// </summary>
    public class CsvReader : TableFileReader, ITableFileReader
    {
        #region ElencyCsvFile
        private ElencyCsvFile _CsvFile = null;
        #endregion

        #region 读取CSV
        public static CsvReader TryLoad(string filePath)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    CsvReader csvReader = new CsvReader(filePath);
                    return csvReader;
                }
                catch (Exception ex)
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="filePath"></param>
        public CsvReader(string filePath)
        { 
            TextReader tr = null;

            try
            {
                ElencyCsvFile csvFile = new ElencyCsvFile();
                csvFile.Populate(filePath, Encoding.UTF8, true, false);
                this._CsvFile = csvFile;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally 
            {
                if (tr != null)
                {
                    tr.Close();
                    tr.Dispose();
                }
            }
        }
        #endregion

        #region GetRowCount
        public override int GetRowCount()
        {
            return this._CsvFile.RecordCount;
        }
        #endregion

        #region GetFieldValues
        public override Dictionary<string, string> GetFieldValues(int rowIndex)
        {
            CsvRecord row = (rowIndex >= 0 && rowIndex < this._CsvFile.RecordCount) ? this._CsvFile.Records[rowIndex] : null;
            if (row != null)
            {
                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                for (int i = 0; i < row.FieldCount; i++)
                {
                    string f = this._CsvFile.Headers[i];
                    string v = row.Fields[i];
                    f2vs.Add(f, v);
                }
                return f2vs;
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region getColumnNameIndex 
        public Dictionary<string, int> GetColumnNameToIndex()
        {
            
            List<string> headers = this._CsvFile.Headers;
            Dictionary<string, int> columnName2Index = new Dictionary<string, int>();
            for (int i = 0; i < headers.Count; i++)
            {
                columnName2Index.Add(headers[i], i);
            }
            return columnName2Index;
        }
        #endregion

    }
}
