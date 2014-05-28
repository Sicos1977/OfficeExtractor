using System;
using System.Runtime.Serialization;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions
{
    /// <summary>
    /// Raised when a method call is invalid for the current object state.
    /// </summary>
    [Serializable]
    public class CFInvalidOperation : CFException
    {
        public CFInvalidOperation() { }

        protected CFInvalidOperation(SerializationInfo info, StreamingContext context) : base(info, context) { }

        public CFInvalidOperation(string message) : base(message, null) { }

        public CFInvalidOperation(string message, Exception innerException) : base(message, innerException) { }
    }
}
