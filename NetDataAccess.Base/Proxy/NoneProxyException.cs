using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.Web
{
    public class NoneProxyException : Exception
    {
        public NoneProxyException(string message)
            : base(message)
        {

        } 
    }
}
