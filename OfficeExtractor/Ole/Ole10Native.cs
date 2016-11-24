using System;
using System.IO;
using OfficeExtractor.Exceptions;
using OpenMcdf;
using System.Text;

/*
   Copyright 2013 - 2016 Kees van Spelde

   Licensed under The Code Project Open License (CPOL) 1.02;
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.codeproject.com/info/cpol10.aspx

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

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
                    var olePackageSize = (int) ole10Native.Size - 4;
                    var olePackageData = new byte[olePackageSize];
                    ole10Native.Read(olePackageData, 4, olePackageSize);
                    var package = new Package(olePackageData);
                    Format = package.Format;
                    FileName = Path.GetFileName(package.FileName);
                    FilePath = package.FilePath;
                    NativeData = package.Data;
                    break;

                case "PBrush":
                    // TODO: Detect in Word if image is visible.
                    var pbBrushSize = (int)ole10Native.Size - 4;
                    var pbBrushData = new byte[pbBrushSize];
                    ole10Native.Read(pbBrushData, 4, pbBrushSize);
                    FileName = Guid.NewGuid() + ".bmp";
                    Format = OleFormat.File;
                    NativeData = pbBrushData;
                    break;

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