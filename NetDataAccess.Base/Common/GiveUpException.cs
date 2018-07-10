using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.Common
{
    public class GiveUpException : Exception
    {
        public GiveUpException()
            : base()
        {

        }
        public GiveUpException(string message)
            : base(message)
        {

        }

        public GiveUpException(string message, Exception innerException)
            : base(message, innerException)
        {

        }
    }
}
