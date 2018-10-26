using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Extended.LunWen.Common
{
    public class BlockedException : Exception
    {
        public BlockedException(string message)
            : base(message)
        {
        }
    }
}
