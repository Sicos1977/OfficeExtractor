using System;
using System.Runtime.Serialization;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions
{
    /// <summary>
    /// Raised when trying to load a Compound File with an invalid, corrupted or mismatched fields (4.1 - specifications).
    /// </summary>
    /// <remarks>
    /// This exception is NOT raised when Compound file has been opened with the NO_VALIDATION_EXCEPTION option.
    /// </remarks>
    [Serializable]
    public class CFCorruptedFileException : CFException
    {
        public CFCorruptedFileException() { }

        protected CFCorruptedFileException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        public CFCorruptedFileException(string message) : base(message, null) { }

        public CFCorruptedFileException(string message, Exception innerException) : base(message, innerException) { }
    }
}
