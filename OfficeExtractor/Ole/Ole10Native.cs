using System;
using System.IO;
using CompoundFileStorage;
using OfficeExtractor.Exceptions;

namespace OfficeExtractor.Ole
{
    /// <summary>
    ///     This class represents an OLE version 2.0 object
    /// </summary>
    /// <remarks>
    ///     See the Microsoft documentation at https://msdn.microsoft.com/en-us/library/dd942280.aspx
    /// </remarks>
    internal class Ole10Native
    {
        #region Properties
        /// <summary>
        ///     This MUST be set to <see cref="OleFormat.Link" /> (0x00000001) or <see cref="OleFormat.File" />
        ///     (0x00000002).
        ///     Otherwise, the ObjectHeader structure is invalid
        /// </summary>
        public OleFormat Format { get; private set; }

        /// <summary>
        ///     This MUST be a LengthPrefixedAnsiString which contain a registered clipboard format name
        /// </summary>
        public string StringFormat { get; private set; }

        /// <summary>
        ///     This MUST be a LengthPrefixedAnsiString structure that contains a display name of the linked
        ///     object or embedded object.
        /// </summary>
        public string AnsiUserType { get; private set; }

        /// <summary>
        ///     AnsiClipboardFormat (variable): This MUST be a ClipboardFormatOrAnsiString structure that contains the
        ///     Clipboard Format of the linked object or embedded object.
        /// </summary>
        public OleClipboardFormat ClipboardFormat { get; private set; }

        /// <summary>
        ///     The filename
        /// </summary>
        public string FileName { get; private set; }

        /// <summary>
        ///     The path to the file before it was embedded
        /// </summary>
        public string FilePath { get; private set; }

        /// <summary>
        ///     The content of the embedded file
        /// </summary>
        public byte[] NativeData { get; private set; }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets all its properties
        /// </summary>
        /// <param name="storage">The OLE version 2.0 object as a <see cref="CFStorage" /></param>
        internal Ole10Native(CFStorage storage)
        {
            if (storage == null)
                throw new ArgumentNullException("storage");

            var ole10Native = storage.GetStream("\x0001Ole10Native");
            var compObj = storage.GetStream("\x0001CompObj");
            var compObjStream = new CompObjStream(compObj);

            AnsiUserType = compObjStream.AnsiUserType;
            StringFormat = compObjStream.StringFormat;
            ClipboardFormat = compObjStream.ClipboardFormat;

            switch (compObjStream.AnsiUserType)
            {
                case "OLE Package":
                    var ole10NativeSize = (int) ole10Native.Size - 4;
                    var data = ole10Native.GetData(4, ref ole10NativeSize);
                    var package = new Package(data);
                    Format = package.Format;
                    FileName = Path.GetFileName(package.FileName);
                    FilePath = package.FilePath;
                    NativeData = package.Data;
                    break;

                case "PBrush":
                case "Pakket":
                    // Ignore
                    break;

                default:
                    throw new OEObjectTypeNotSupported("Unsupported OleNative AnsiUserType '" +
                                                        compObjStream.AnsiUserType + "' found");
            }
        }
        #endregion
    }
}