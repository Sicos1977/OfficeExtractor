using System;
using System.Runtime.Serialization;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions
{
    /// <summary>
    ///     Raised when opening a file with invalid header or not supported COM/OLE Structured storage version.
    /// </summary>
    [Serializable]
    public class CFFileFormatException : CFException
    {
        public CFFileFormatException()
        {
        }

        protected CFFileFormatException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }

        public CFFileFormatException(string message) : base(message, null)
        {
        }

        public CFFileFormatException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}