namespace OfficeExtractor.Ole
{
    #region Enum OleObjectFormat
    /// <summary>
    /// Type OLE version 1.0 and 2.0 object type
    /// </summary>
    internal enum OleObjectFormat
    {
        /// <summary>
        /// The format is NOT set
        /// </summary>
        NotSet = 0x00000000,

        /// <summary>
        /// The embedded object is a link
        /// </summary>
        Link = 0x00000001,

        /// <summary>
        /// The embedded object is a file
        /// </summary>
        File = 0x00000002,

        /// <summary>
        /// The embedded object is a presentation (e.g. a image)
        /// </summary>
        Presentation = 0x00000005
    }
    #endregion

    #region Enum OleClipboardFormat
    /// <summary>
    /// The OLE version 1.0 and 2.0 clipboard formats
    /// </summary>
    internal enum OleClipboardFormat
    {
        /// <summary>
        /// The format is a registered clipboard format
        /// </summary>
        /// <remarks>
        /// The format is set into the <see cref="ObjectV10.StringFormatData"/> field
        /// </remarks>
        Registered = 0x00000000,

        // ReSharper disable InconsistentNaming
        /// <summary>
        /// Bitmap16 Object structure 
        /// </summary>
        CF_BITMAP = 0x00000002,

        /// <summary>
        /// 
        /// </summary>
        CF_METAFILEPICT = 0x00000003,

        /// <summary>
        /// DeviceIndependentBitmap Object structure
        /// </summary>
        CF_DIB = 0x00000008,

        /// <summary>
        /// Enhanced Metafile
        /// </summary>
        CF_ENHMETAFILE = 0x0000000E
        // ReSharper restore InconsistentNaming
    }
    #endregion
}
