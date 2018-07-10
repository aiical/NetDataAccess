using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.DataTransform.Address
{
    public class XZQHMap
    {
        private Dictionary<string, XZQHArea> _AreaMap = new Dictionary<string, XZQHArea>();
        public Dictionary<string, XZQHArea> AreaMap
        {
            get
            {
                return this._AreaMap;
            }
            set
            {
                this._AreaMap = value;
            }
        }

        private List<string> _RootAreaCodes = null;
        public List<string> RootAreaCodes
        {
            get
            {
                return this._RootAreaCodes;
            }
            set
            {
                this._RootAreaCodes = value;
            }
        }

        public XZQHArea GetArea(string code)
        {
            if (this.AreaMap.ContainsKey(code))
            {
                return this.AreaMap[code];
            }
            else
            {
                return null;
            }
        }

    }
}
