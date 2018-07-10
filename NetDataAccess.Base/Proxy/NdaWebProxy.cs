using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace NetDataAccess.Base.Proxy
{
    public class NdaWebProxy : WebProxy
    {
        private int _Index = 0;
        public int Index
        {
            get
            {
                return _Index;
            }
            set 
            {
                _Index = value;
            }
        }

        public NdaWebProxy(int index)
            : base()
        {
            this.Index = index;
        }
    }
}
