using OfficeExtractor.Helpers;

namespace OfficeExtractor;

internal abstract class OfficeBase
{
    #region Properties
    /// <summary>
    ///     Returns a reference to the Extraction class when it already exists or creates a new one
    ///     when it doesn't
    /// </summary>
    protected Extraction Extraction
    {
        get
        {
            _extraction ??= new Extraction();
            return _extraction;
        }
    }
    #endregion

    #region Fields

    //TODO: once NET10/CS14 is here, use the new 'field' keyword and remove the backing field
    private Extraction _extraction;

    #endregion

    protected void HandleException(System.Exception exception, string extractionFile, bool shallThrow = true)
    {
        Logger.WriteToLog($"An error occurred while extracting an embedded object from the {extractionFile} document: {exception}");
        // It may well be that a document contains more than one embedded objects
        // and we want to extract as much as possible
        if (shallThrow)
            throw exception;
    }
}
