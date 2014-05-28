using System;
using System.Runtime.Serialization;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions
{
    /// <summary>
    /// The base exception for all the exceptions that are raised in the CompoundFileStorage namespace
    /// </summary>
    [Serializable]
    public class CFException : Exception
    {
        public CFException() { }

        protected CFException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        public CFException(string message) : base(message, null) { }

        public CFException(string message, Exception innerException) : base(message, innerException) { }

    }
}
