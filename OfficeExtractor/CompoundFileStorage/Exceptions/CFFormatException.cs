using System;
using System.Runtime.Serialization;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions
{
    /// <summary>
    ///     Raised when a file with an invalid header or a NON supported COM/OLE structured storage version is opened.
    /// </summary>
    [Serializable]
    public class CFFormatException : CFException
    {
        public CFFormatException()
        {
        }

        protected CFFormatException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }

        public CFFormatException(string message) : base(message, null)
        {
        }

        public CFFormatException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}