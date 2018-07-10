using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.Reader
{
    public class TableFileReader : ITableFileReader
    {
        public List<string> GetColumnValues(string columnName)
        {
            List<string> values = new List<string>();
            int rowCount = this.GetRowCount();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> row = this.GetFieldValues(i);
                values.Add(row[columnName]);
            }
            return values;
        }
        public List<Dictionary<string, string>> GetColumnValues(string[] columnNames)
        {
            List<Dictionary<string, string>> valuesLists = new List<Dictionary<string, string>>();
            int rowCount = this.GetRowCount();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> row = this.GetFieldValues(i);

                Dictionary<string, string> values = new Dictionary<string, string>();
                foreach (string columnName in columnNames)
                {
                    values.Add(columnName, row[columnName]);
                }
                valuesLists.Add(values);
            }
            return valuesLists;
        }

        public virtual Dictionary<string, string> GetFieldValues(int rowIndex)
        {
            return null;
        }

        public virtual int GetRowCount()
        {
            return 0;
        }
    }
}
