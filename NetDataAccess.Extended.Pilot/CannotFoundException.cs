using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Extended.Pilot
{
    public class CannotFoundException:Exception
    {
        public CannotFoundException(string message)
            : base(message)
        {
        }
    }
}
