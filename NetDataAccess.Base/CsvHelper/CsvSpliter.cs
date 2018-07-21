using NetDataAccess.Base.Reader;
using NetDataAccess.Base.Writer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.CsvHelper
{
    public class CsvSpliter
    {
        private string _SourceFilePath = "";
        public string SourceFilePath
        {
            get
            {
                return this._SourceFilePath;
            }
        }

        private CsvReader _CR = null;
        public CsvReader CR
        {
            get
            {
                return this._CR;
            }
        }

        public int Init(string sourceFilePath)
        {
            this._SourceFilePath = sourceFilePath;
            _CR = new CsvReader(sourceFilePath);
            return CR.GetRowCount();
        }

        public void GetPart(string destFilePath, int fromRowIndex, int rowCount)
        {
            Dictionary<string, int> columnName2Index = CR.GetColumnNameToIndex();
            CsvWriter cw = new CsvWriter(destFilePath, columnName2Index);
            int rightRowCount = CR.GetRowCount() - fromRowIndex;
            if (rightRowCount < 0)
            {
                throw new Exception("获取csv部分数据时, 起始行超出总行数");
            }
            else if (rightRowCount < rowCount)
            {
                rowCount = rightRowCount;
            }

            int toRowIndex = rowCount + fromRowIndex - 1;
            for (int i = fromRowIndex; i <= toRowIndex; i++)
            {
                Dictionary<string, string> row = CR.GetFieldValues(i);
                cw.AddRow(row);
            }
            cw.SaveToDisk();
        }
    }
}
