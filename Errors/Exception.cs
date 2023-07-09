using System;

namespace CADShark.Common.SolidWorks.Errors
{
    public class ConnectionException : Exception
    {
        public ConnectionException(string message) : base(message)
        {
        }
    }
}
