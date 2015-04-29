using System;
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
        /// Returns the format for the data that is stored inside the Ole10Native stream.
        /// </summary>
        public OleFormat Format { get; private set; }

        /// <summary>
        ///     This MUST be a LengthPrefixedUnicodeString which contain a registered clipboard format name
        /// </summary>
        public string StringFormatData { get; private set; }

        /// <summary>
        ///     Identifies the <see cref="NativeData" /> when this file is a Clipboard object
        /// </summary>
        public OleClipboardFormat ClipboardFormat { get; private set; }

        /// <summary>
        ///     This MUST be a LengthPrefixedAnsiString that contains a value identifying the creating application.
        ///     The value is mapped to the creating application in an implementation-specific manner
        /// </summary>
        public string ClassName { get; private set; }

        /// <summary>
        ///     The name of the file
        /// </summary>
        public string FileName { get; private set; }

        /// <summary>
        ///     The original location of the file (before it was embedded)
        /// </summary>
        public string FilePath { get; private set; }

        /// <summary>
        ///     Returns the embedded file when <see cref="Format"/> is set to <see cref="OleFormat.File"/>, 
        ///     otherwise this array will be empty
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

            //var ole = storage.GetStream("\x0001Ole");
            //var oleStream = new OleStream(ole);
            var compObj = storage.GetStream("\x0001CompObj");
            var compObjStream = new CompObjStream(compObj);

            switch (compObjStream.StringFormat)
            {
                case "OLE Package":
                    break;

                default:
                    throw new OEObjectTypeNotSupported("Unsupported OleNative stringformat '" +
                                                       compObjStream.StringFormat + "' found");
            }

            if (compObjStream.StringFormat == "OLE Package")
            {
                var package = new Package(ole10Native);
                FileName = package.FileName;
                FilePath = package.FilePath;
                NativeData = package.Data;
            }
        }
        #endregion
    }
}