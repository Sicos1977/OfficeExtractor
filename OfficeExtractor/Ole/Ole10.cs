using System;
using System.IO;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;

//
// Ole10.cs
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
    ///     This class represents an OLE version 1.0 object
    /// </summary>
    /// <remarks>
    ///     See the Microsoft documentation at https://msdn.microsoft.com/en-us/library/dd942076.aspx
    /// </remarks>
    internal class Ole10
    {
        #region Properties
        /// <summary>
        ///     OLEVersion (4 bytes): This can be set to any arbitrary value and MUST be ignored on processing
        /// </summary>
        public uint Version { get; private set; }

        /// <summary>
        ///     This MUST be set to <see cref="OleFormat.Link" /> (0x00000001) or <see cref="OleFormat.File" />
        ///     (0x00000002).
        ///     Otherwise, the ObjectHeader structure is invalid
        /// </summary>
        /// <remarks>
        ///     0x00000001 = The ObjectHeader structure MUST be followed by a LinkedObject structure.
        ///     0x00000002 = The ObjectHeader structure MUST be followed by an EmbeddedObject structure.
        /// </remarks>
        public OleFormat Format { get; private set; }

        /// <summary>
        ///     This MUST be a LengthPrefixedAnsiString that contains a value identifying the creating application.
        ///     The value is mapped to the creating application in an implementation-specific manner
        /// </summary>
        public string ClassName { get; private set; }

        /// <summary>
        ///     This MUST be set to the width of the presentation object. If the ClassName field of the Header
        ///     is set  to the case-sensitive value "METAFILEPICT", this MUST be a MetaFilePresentationDataWidth.
        ///     If the ClassName field of the Header is set to either the case-sensitive value "BITMAP" or the
        ///     case-sensitive value "DIB", this MUST be a DIBPresentationDataWidth.
        /// </summary>
        /// <remarks>
        ///     If <see cref="Format" /> is set to <see cref="OleFormat.File" /> then this property is 0
        /// </remarks>
        public long Width { get; private set; }

        /// <summary>
        ///     This MUST be set to the width of the presentation object. If the ClassName field of the Header is
        ///     set to the case-sensitive value "METAFILEPICT", this MUST be a MetaFilePresentationDataWidth.
        ///     If the ClassName field of the Header is set to either the case-sensitive value "BITMAP" or the
        ///     case-sensitive value "DIB", this MUST be a DIBPresentationDataWidth.
        /// </summary>
        /// <remarks>
        ///     If <see cref="Format" /> is set to <see cref="OleFormat.File" /> then this property is 0
        /// </remarks>
        public long Height { get; private set; }

        /// <summary>
        ///     This MUST be a LengthPrefixedAnsiString or a LengthPrefixedUnicodeString, either of which contain a
        ///     registered clipboard format name
        /// </summary>
        /// <remarks>
        ///     Only set when <see cref="Format" /> is set to <see cref="OleFormat.Presentation" /> and the
        ///     <see cref="ClassName" /> does not contain "METAFILEPICT", "BITMAP" or "DIB"
        /// </remarks>
        public string StringFormat { get; private set; }

        /// <summary>
        ///     Identifies the <see cref="NativeData" /> when this file is a Clipboard object
        /// </summary>
        /// <remarks>
        ///     Only set when <see cref="Format" /> is set to <see cref="OleFormat.Presentation" /> and the
        ///     <see cref="ClassName" /> does not contain "METAFILEPICT", "BITMAP" or "DIB"
        /// </remarks>
        public OleClipboardFormat ClipboardFormat { get; private set; }

        /// <summary>
        ///     If the ObjectHeader structure is contained by an EmbeddedObject,
        ///     the TopicName field SHOULD contain an empty string and MUST be ignored on processing
        /// </summary>
        /// <remarks>
        ///     If <see cref="Format" /> is set to <see cref="OleFormat.File" /> then this property is null
        /// </remarks>
        public string TopicName { get; private set; }

        /// <summary>
        ///     ItemName (variable): This MUST be a LengthPrefixedAnsiString.
        ///     If the ObjectHeader structure is contained by an EmbeddedObject structure,
        ///     the ItemName field SHOULD contain an empty string and MUST be ignored on processing.
        ///     If the ObjectHeader structure is contained by a LinkedObject structure,the ItemName field
        ///     MUST contain a string that is used by the application or higher-level protocol to identify
        ///     the item within the file to which is being linked. The format and meaning of the ItemName
        ///     string is specific to the creating application and MUST be treated by other parties as an
        ///     opaque string when processing this data structure. An example of such an item is an
        ///     individual cell within a spreadsheet application.
        /// </summary>
        /// <remarks>
        ///     If <see cref="Format" /> is set to <see cref="OleFormat.File" /> then this property is empty
        /// </remarks>
        public string ItemName { get; private set; }

        /// <summary>
        ///     If the <see cref="TopicName" /> field of the ObjectHeader structure contains a path that starts
        ///     with a drive letter and if the drive letter is for a remote drive, the NetworkName field MUST
        ///     contain the path name of the linked file in the Universal Naming Convention (UNC) format.
        /// </summary>
        /// <remarks>
        ///     If <see cref="Format" /> is set to <see cref="OleFormat.File" /> then this property is empty
        /// </remarks>
        public string NetworkName { get; private set; }

        /// <summary>
        ///     This field contains an implementation-specific hint supplied by the application or higher-level
        ///     protocol responsible for creating the data structure. The hint MAY be ignored on processing of
        ///     this data structure
        /// </summary>
        /// <remarks>
        ///     If <see cref="Format" /> is set to <see cref="OleFormat.File" /> then this property is empty
        /// </remarks>
        public uint LinkUpdateOptions { get; private set; }

        /// <summary>
        ///     Data that is required to display the linked or embedded object within the container application.
        /// </summary>
        public byte[] PresentationData { get; private set; }

        /// <summary>
        ///     The data that constitutes the state of an embedded object. The only entity that can create
        ///     and process the data is the creating application.
        /// </summary>
        public byte[] NativeData { get; private set; }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets all its properties
        /// </summary>
        /// <param name="inputStream">The OLE version 1.0 object as a stream</param>
        /// <exception cref="ArgumentNullException">Raised when <paramref name="inputStream" /> is <c>null</c></exception>
        public Ole10(Stream inputStream)
        {
            if (inputStream == null)
                throw new ArgumentNullException(nameof(inputStream));

            inputStream.Position = 0;

            using (var binaryReader = new BinaryReader(inputStream))
                ParseOle(binaryReader);
        }

        /// <summary>
        ///     Creates this object and sets all its properties
        /// </summary>
        /// <param name="reader">The OLE version 1.0 object as a stream</param>
        /// <exception cref="ArgumentNullException">Raised when <paramref name="reader" /> is <c>null</c></exception>
        internal Ole10(BinaryReader reader)
        {
            if (reader == null)
                throw new ArgumentNullException(nameof(reader));

            ParseOle(reader);
        }
        #endregion

        #region ParseOle
        /// <summary>
        ///     Parses the stream and sets all the OLE properties
        /// </summary>
        /// <param name="binaryReader"></param>
        private void ParseOle(BinaryReader binaryReader)
        {
            Version = binaryReader.ReadUInt32();

            var format = binaryReader.ReadUInt32(); // FormatID
            try
            {
                Format = (OleFormat)format;
            }
            catch (Exception)
            {
                throw new OEObjectTypeNotSupported(
                    "Invalid OLE version 1.0 format, expected 0x00000000, 0x00000002 or 0x00000005");
            }

            if (Format != OleFormat.NotSet)
                ClassName = Strings.Read4ByteLengthPrefixedAnsiString(binaryReader);

            switch (Format)
            {
                case OleFormat.Link:
                    ParseObjectHeader(binaryReader);
                    ParseLinkedObject(binaryReader);
                    break;

                case OleFormat.File:
                    ParseObjectHeader(binaryReader);
                    ParseEmbeddedObject(binaryReader);
                    break;

                case OleFormat.Presentation:
                    switch (ClassName)
                    {
                        // MetaFilePresentationObject
                        case "METAFILEPICT":
                        case "BITMAP":
                        case "DIB":
                            ParseStandardPresentationObject(binaryReader);
                            break;

                        default:
                            ParseGenericPresentationObject(binaryReader);
                            break;
                    }

                    break;
            }
        }
        #endregion

        #region ParseStandardPresentationObject
        /// <summary>
        ///     Parses the standard presentation object when the <see cref="Format" /> is set to
        ///     <see cref="OleFormat.Presentation" /> and
        ///     the <see cref="ClassName" /> is set to "METAFILEPICT", "BITMAP" or "DIB"
        /// </summary>
        /// <param name="binaryReader"></param>
        private void ParseStandardPresentationObject(BinaryReader binaryReader)
        {
            Width = binaryReader.ReadUInt32();
            Height = binaryReader.ReadUInt32();

            switch (ClassName)
            {
                // MetaFilePresentationObject
                case "METAFILEPICT":
                    {
                        var size = binaryReader.ReadUInt32();

                        // PresentationDataSize (4 bytes): This MUST be an unsigned long integer set to the sum of the size,
                        // in bytes, of the PresentationData field and the number 8. If this field contains the value 8, 
                        // the PresentationData field MUST NOT be present.
                        if (size == 8)
                            return;

                        // Reserved1 (2 bytes): Reserved. This can be set to any arbitrary value and MUST be ignored on processing.
                        binaryReader.ReadUInt16();
                        // Reserved2 (2 bytes): Reserved. This can be set to any arbitrary value and MUST be ignored on processing.
                        binaryReader.ReadUInt16();
                        // Reserved3 (2 bytes): Reserved. This can be set to any arbitrary value and MUST be ignored on processing.
                        binaryReader.ReadUInt16();
                        // Reserved4 (2 bytes): Reserved. This can be set to any arbitrary value and MUST be ignored on processing.
                        binaryReader.ReadUInt16();

                        // This MUST be an array of bytes that contain a Windows metafile (as specified in [MS-WMF]).
                        if (size - 8 > 0)
                            PresentationData = binaryReader.ReadBytes((int)size - 8);

                        break;
                    }

                // BitmapPresentationObject
                case "BITMAP":
                case "DIB":
                    {
                        var size = binaryReader.ReadUInt32();

                        // PresentationDataSize (4 bytes): This MUST be an unsigned long integer set to the size, 
                        // in bytes, of the Bitmap or DIB field. If this field has the value 0, the Bitmap or DIB field MUST 
                        // NOT be present.
                        if (size == 0)
                            return;

                        PresentationData = binaryReader.ReadBytes((int)size);
                        break;
                    }

                default:
                    throw new OEObjectTypeNotSupported(
                        "Unsupported classname '" + ClassName + "' found, expected METAFILEPICT, BITMAP or DIB");
            }
        }
        #endregion

        #region ParseGenericPresentationObject
        /// <summary>
        ///     Parses the generic presentation object when the <see cref="Format" /> is set to
        ///     <see cref="OleFormat.Presentation" /> and
        ///     the <see cref="ClassName" /> is <b>NOT</b> set to "METAFILEPICT", "BITMAP" or "DIB"
        /// </summary>
        /// <param name="binaryReader"></param>
        private void ParseGenericPresentationObject(BinaryReader binaryReader)
        {
            ClipboardFormat = (OleClipboardFormat) binaryReader.ReadUInt32();

            switch (ClipboardFormat)
            {
                // RegisteredClipboardFormatPresentationObject
                case OleClipboardFormat.Registered:
                {
                    // This MUST be set to the size, in bytes, of the StringFormatData field.
                    // ReSharper disable once UnusedVariable
                    var stringFormatDataSize = binaryReader.ReadUInt32();
                    StringFormat = Strings.Read4ByteLengthPrefixedAnsiString(binaryReader);
                    var size = binaryReader.ReadUInt32();
                    NativeData = binaryReader.ReadBytes((int) size);
                    break;
                }

                case OleClipboardFormat.CF_BITMAP:
                case OleClipboardFormat.CF_DIB:
                case OleClipboardFormat.CF_ENHMETAFILE:
                case OleClipboardFormat.CF_METAFILEPICT:
                {
                    var size = binaryReader.ReadUInt32();
                    NativeData = binaryReader.ReadBytes((int) size);
                    break;
                }

                default:
                    throw new OEObjectTypeNotSupported(
                        "A not supported clipboardformat has been found, only CF_BITMAP, CF_DIB, CF_ENHMETAFILE and CF+METAFILEPICT are supported");
            }
        }
        #endregion

        #region ParseObjectHeader
        /// <summary>
        ///     Parses the stream when the <see cref="Format" /> is set to <see cref="OleFormat.File" />
        ///     or set to <see cref="OleFormat.Link" />
        /// </summary>
        /// <param name="binaryReader"></param>
        private void ParseObjectHeader(BinaryReader binaryReader)
        {
            TopicName = Strings.Read4ByteLengthPrefixedAnsiString(binaryReader);
            ItemName = Strings.Read4ByteLengthPrefixedAnsiString(binaryReader);
        }
        #endregion

        #region GetAndSetPresentationObject
        /// <summary>
        ///     A <see cref="ParseEmbeddedObject" /> or <see cref="ParseLinkedObject" /> always contain
        ///     a presentation object, this is used to present the embedded link or object in the host
        ///     application. This method will extract this data.
        /// </summary>
        /// <param name="binaryReader"></param>
        public void GetAndSetPresentationObject(BinaryReader binaryReader)
        {
            var po = new Ole10(binaryReader);
            Width = po.Width;
            Height = po.Height;
            StringFormat = po.StringFormat;
            ClipboardFormat = po.ClipboardFormat;
            PresentationData = po.PresentationData;
        }
        #endregion

        #region ParseEmbeddedObject
        /// <summary>
        ///     Parses the stream when it is an <see cref="OleFormat.File" />
        /// </summary>
        /// <param name="binaryReader"></param>
        private void ParseEmbeddedObject(BinaryReader binaryReader)
        {
            var nativeDataSize = binaryReader.ReadUInt32();
            NativeData = binaryReader.ReadBytes((int) nativeDataSize);
            GetAndSetPresentationObject(binaryReader);
        }
        #endregion

        #region ParseLinkedObject
        /// <summary>
        ///     Parses the stream when it is an <see cref="OleFormat.Link" />
        /// </summary>
        /// <param name="binaryReader"></param>
        private void ParseLinkedObject(BinaryReader binaryReader)
        {
            NetworkName = Strings.Read4ByteLengthPrefixedAnsiString(binaryReader);
            TopicName = Strings.Read4ByteLengthPrefixedAnsiString(binaryReader);
            // Reserved (4 bytes)
            binaryReader.ReadUInt32();
            LinkUpdateOptions = binaryReader.ReadUInt32();
            GetAndSetPresentationObject(binaryReader);
        }
        #endregion
    }
}