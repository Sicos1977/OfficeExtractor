using System;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.Exceptions
{
    /// <summary>
    ///     Raised when there is an Excel configuration problem
    /// </summary>
    public class OEExcelConfiguration : Exception
    {
        internal OEExcelConfiguration() { }

        internal OEExcelConfiguration(string message) : base(message) { }

        internal OEExcelConfiguration(string message, Exception inner) : base(message, inner) { }
    }
}
