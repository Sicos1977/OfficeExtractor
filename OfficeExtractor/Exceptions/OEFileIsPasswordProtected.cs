using System;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.Exceptions
{
    /// <summary>
    /// Raised when the Microsoft Office file is password protected
    /// </summary>
    public class OEFileIsPasswordProtected : Exception
    {
        public OEFileIsPasswordProtected() { }

        public OEFileIsPasswordProtected(string message) : base(message) { }

        public OEFileIsPasswordProtected(string message, Exception inner) : base(message, inner) { }
    }
}
