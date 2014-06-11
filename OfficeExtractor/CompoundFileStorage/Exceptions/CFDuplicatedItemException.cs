using System;
using System.Runtime.Serialization;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions
{
    /// <summary>
    ///     Raised when trying to add a duplicated CFItem.
    /// </summary>
    /// <remarks>
    ///     Items are compared by name as indicated by specs. Two items with the same name CANNOT be added within
    ///     the same storage or sub-storage.
    /// </remarks>
    [Serializable]
    public class CFDuplicatedItemException : CFException
    {
        public CFDuplicatedItemException()
        {
        }

        protected CFDuplicatedItemException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }

        public CFDuplicatedItemException(string message) : base(message, null)
        {
        }

        public CFDuplicatedItemException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}