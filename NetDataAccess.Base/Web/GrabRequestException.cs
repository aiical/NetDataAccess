using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.Web
{
    public class GrabRequestException : Exception
    {
        public GrabRequestException(string message, Exception innerException)
            : base(message, innerException)
        {

        }
        public GrabRequestException(string message)
            : base(message)
        {

        }
    }
}
