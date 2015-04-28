using System;

namespace OfficeExtractor.Exceptions
{
    /// <summary>
    ///     Raised when an unsupported embedded object has been found
    /// </summary>
    public class OEObjectTypeNotSupported : Exception
    {
        public OEObjectTypeNotSupported()
        {
        }

        public OEObjectTypeNotSupported(string message) : base(message)
        {
        }

        public OEObjectTypeNotSupported(string message, Exception inner) : base(message, inner)
        {
        }
    }
}