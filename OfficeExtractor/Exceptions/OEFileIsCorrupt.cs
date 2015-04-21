using System;

namespace OfficeExtractor.Exceptions
{
    /// <summary>
    ///     Raised when the Microsoft Office file is corrupt
    /// </summary>
    public class OEFileIsCorrupt : Exception
    {
        public OEFileIsCorrupt()
        {
        }

        public OEFileIsCorrupt(string message) : base(message)
        {
        }

        public OEFileIsCorrupt(string message, Exception inner) : base(message, inner)
        {
        }
    }
}