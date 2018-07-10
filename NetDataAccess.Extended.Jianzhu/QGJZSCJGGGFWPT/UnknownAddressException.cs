using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Extended.Jianzhu.QGJZSCJGGGFWPT
{
    public class UnknownAddressException : Exception
    {
        public UnknownAddressException(string message)
            : base(message)
        {
        }
    }
}
