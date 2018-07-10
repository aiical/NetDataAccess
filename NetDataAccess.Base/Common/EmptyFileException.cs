using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.Common
{
    public class EmptyFileException : Exception
    {
        public EmptyFileException()
            : base()
        {

        }
        public EmptyFileException(string message)
            : base(message)
        {

        }

        public EmptyFileException(string message, Exception innerException)
            : base(message, innerException)
        {

        }
    }
}
