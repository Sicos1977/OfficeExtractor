using System;

namespace OfficeExtractor.Exceptions
{
    /// <summary>
    ///     Raised when the Microsoft Office file type is not supported
    /// </summary>
    public class OEFileTypeNotSupported : Exception
    {
        public OEFileTypeNotSupported()
        {
        }

        public OEFileTypeNotSupported(string message) : base(message)
        {
        }

        public OEFileTypeNotSupported(string message, Exception inner) : base(message, inner)
        {
        }
    }
}