using System;
using System.Runtime.Serialization;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions
{
    /// <summary>
    /// Raised when a data setter/getter method is invoked on a stream or storage object after the 
    /// disposal of the owner compound file object.
    /// </summary>
    [Serializable]
    public class CFDisposedException : CFException
    {
        public CFDisposedException() { }

        protected CFDisposedException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        public CFDisposedException(string message) : base(message, null) { }

        public CFDisposedException(string message, Exception innerException) : base(message, innerException) { }
    }
}
