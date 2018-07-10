using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.Reader
{
    public interface ITableFileReader
    { 
        Dictionary<string, string> GetFieldValues(int rowIndex);

        int GetRowCount();

        List<string> GetColumnValues(string columnName);
    }


}
