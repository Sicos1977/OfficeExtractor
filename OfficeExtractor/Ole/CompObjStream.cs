using System.IO;
using OfficeExtractor.Helpers;
using OpenMcdf;

//
// CompObjStream.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2020 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

namespace OfficeExtractor.Ole
{
    /// <summary>
    /// The CompObjStream structure is contained inside of an OLE Compound File Stream. The OLE Compound File Stream 
    /// has the name "\1CompObj". The CompObjStream structure specifies the Clipboard Format and the display name of 
    /// the linked object or embedded object.
    /// </summary>
    internal class CompObjStream
    {
        #region Properties
        /// <summary>
        ///     This MUST be a LengthPrefixedAnsiString structure that contains a display name of the linked
        ///     object or embedded object.
        /// </summary>
        public string AnsiUserType { get; private set; }

        /// <summary>
        ///     This MUST be a LengthPrefixedAnsiString which contain a registered clipboard format name
        /// </summary>
        public string StringFormat { get; private set; }

        /// <summary>
        ///     AnsiClipboardFormat (variable): This MUST be a ClipboardFormatOrAnsiString structure that contains the
        ///     Clipboard Format of the linked object or embedded object.
        /// </summary>
        public OleClipboardFormat ClipboardFormat { get; private set; }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets all its properties
        /// </summary>
        /// <param name="stream">The Compound File Storage CompObj <see cref="CFStream" /></param>
        internal CompObjStream(CFStream stream)
        {
            using (var memoryStream = new MemoryStream(stream.GetData()))
            using (var binaryReader = new BinaryReader(memoryStream))
            {
                // Reserved1 (4 bytes): This can be set to any arbitrary value and MUST be ignored on processing.
                // Version (4 bytes): This can be set to any arbitrary value and MUST be ignored on processing.
                // Reserved2 (20 bytes): This can be set to any arbitrary value and MUST be ignored on processing.
                // Skip the first 28 bytes, this is the CompObjHeader
                binaryReader.ReadBytes(28);

                // This MUST be a LengthPrefixedAnsiString structure that contains a display name of the linked 
                // object or embedded object. 
                AnsiUserType = Strings.Read4ByteLengthPrefixedAnsiString(binaryReader);

                // MarkerOrLength (4 bytes): If this is set to 0x00000000, the FormatOrAnsiString field MUST NOT 
                // be present. If this field is set to 0xFFFFFFFF or 0xFFFFFFFE, the FormatOrAnsiString field MUST 
                // be 4 bytes in size and MUST contain a standard clipboard format identifier. 
                // If this set to a value other than 0x00000000, the FormatOrAnsiString field MUST be set to a 
                // null-terminated ANSI string containing the name of a registered clipboard format and the 
                // MarkerOrLength field MUST be set to the number of ANSI characters in the FormatOrAnsiString field, 
                // including the terminating null character.
                var markerOrLength = binaryReader.ReadUInt32();

                switch (markerOrLength)
                {
                    case 0x00000000:
                        // Skip
                        break;

                    case 0xFFFFFFFF:
                    case 0xFFFFFFFE:
                        ClipboardFormat = (OleClipboardFormat) binaryReader.ReadUInt32();
                        break;

                    default:
                        binaryReader.BaseStream.Position -= 4;
                        StringFormat = Strings.Read4ByteLengthPrefixedAnsiString(binaryReader);
                        break;
                }


                // Reserved1 (variable): If present, this MUST be a LengthPrefixedAnsiString structure. If the Length 
                // field of the LengthPrefixedAnsiString contains a value of 0 or a value that is greater than 0x00000028, 
                // the remaining fields of the structure starting with the String field of the LengthPrefixedAnsiString 
                // MUST be ignored on processing. 
                var reserved1Length = binaryReader.ReadUInt32();
                if (reserved1Length <= 0x00000028)
                {
                    binaryReader.BaseStream.Position -= 4;
                    // ReSharper disable once UnusedVariable
                    var reserved1 = Strings.Read4ByteLengthPrefixedAnsiString(binaryReader);
                }

                // UnicodeMarker (variable): If this field is present and is NOT set to 0x71B239F4, the remaining fields 
                // of the structure MUST be ignored on processing.
                var unicodeMarker = binaryReader.ReadUInt32();
                if (unicodeMarker == 0x71B239F4)
                {
                    markerOrLength = binaryReader.ReadUInt32();

                    switch (markerOrLength)
                    {
                        case 0x00000000:
                            // Skip
                            break;

                        case 0xFFFFFFFF:
                        case 0xFFFFFFFE:
                            ClipboardFormat = (OleClipboardFormat) binaryReader.ReadUInt32();
                            break;

                        default:
                            binaryReader.BaseStream.Position -= 4;
                            StringFormat = Strings.Read4ByteLengthPrefixedAnsiString(binaryReader);
                            break;
                    }
                }
            }
        }
        #endregion
    }
}