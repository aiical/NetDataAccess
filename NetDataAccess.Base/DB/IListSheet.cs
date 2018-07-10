using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.DB
{
    public interface IListSheet
    {
        int RowCount { get; }

        List<string> PageUrlList { get; }

        List<string> PageNameList { get; }

        List<string> PageCookieList { get; }

        List<bool> GiveUpList { get; }

        void SetGiveUp(int pageIndex, string pageUrl, string errorMsg);

        void InitDetailPageInfo();

        int GetSucceedListPageIndex();

        void CopyToListSheet(List<Dictionary<string, string>> allF2vs, Dictionary<string, int> columnToIndexs, List<string> sysColumnList);

        int GetListDBRowCount();
        
        Dictionary<string,string> GetRow(int rowIndex);

        void AddListRow(Dictionary<string, string> f2vs);

        void AddColumns(List<string> newColumns);

        void Close();
    }
}
