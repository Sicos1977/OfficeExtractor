using System;
using System.Runtime.Serialization;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions
{
    /// <summary>
    ///     Raised when a named stream or a storage object is not found in a parent storage.
    /// </summary>
    [Serializable]
    public class CFItemNotFound : CFException
    {
        protected CFItemNotFound(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }

        public CFItemNotFound() : base("Entry not found")
        {
        }

        public CFItemNotFound(string message) : base(message, null)
        {
        }

        public CFItemNotFound(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}