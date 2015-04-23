using System;
using System.IO;
using System.Text;
using OfficeExtractor.Exceptions;

namespace OfficeExtractor.Ole
{
    #region Enum Ole10ObjectFormat
    /// <summary>
    /// Type OLE version 1.0 object type
    /// </summary>
    internal enum Ole10ObjectFormat
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

    #region Enum ClipboardFormat
    /// <summary>
    /// The standard clipboard formsts
    /// </summary>
    internal enum ClipboardFormat
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

    /// <summary>
    /// This class represents an OLE version 1.0 object
    /// </summary>
    /// <remarks>
    /// See the Microsoft documentation at https://msdn.microsoft.com/en-us/library/dd942076.aspx
    /// </remarks>
    internal class ObjectV10
    {
        #region Properties
        /// <summary>
        /// OLEVersion (4 bytes): This can be set to any arbitrary value and MUST be ignored on processing
        /// </summary>
        public UInt32 Version { get; private set; }

        /// <summary>
        /// This MUST be set to <see cref="Ole10ObjectFormat.Link"/> (0x00000001) or <see cref="Ole10ObjectFormat.File"/> (0x00000002). 
        /// Otherwise, the ObjectHeader structure is invalid
        /// </summary>
        /// <remarks>
        /// 0x00000001 = The ObjectHeader structure MUST be followed by a LinkedObject structure.
        /// 0x00000002 = The ObjectHeader structure MUST be followed by an EmbeddedObject structure.
        /// </remarks>
        public Ole10ObjectFormat Format { get; private set; }

        /// <summary>
        /// This MUST be a LengthPrefixedAnsiString that contains a value identifying the creating application. 
        /// The value is mapped to the creating application in an implementation-specific manner
        /// </summary>
        public string ClassName { get; private set; }

        /// <summary>
        /// This MUST be set to the width of the presentation object. If the ClassName field of the Header 
        /// is set  to the case-sensitive value "METAFILEPICT", this MUST be a MetaFilePresentationDataWidth. 
        /// If the ClassName field of the Header is set to either the case-sensitive value "BITMAP" or the 
        /// case-sensitive value "DIB", this MUST be a DIBPresentationDataWidth.
        /// </summary>
        /// <remarks>
        /// If <see cref="Format"/> is set to <see cref="Ole10ObjectFormat.File"/> then this property is empty
        /// </remarks>
        public long Width { get; private set; }

        /// <summary>
        /// This MUST be set to the width of the presentation object. If the ClassName field of the Header is 
        /// set to the case-sensitive value "METAFILEPICT", this MUST be a MetaFilePresentationDataWidth. 
        /// If the ClassName field of the Header is set to either the case-sensitive value "BITMAP" or the 
        /// case-sensitive value "DIB", this MUST be a DIBPresentationDataWidth.
        /// </summary>
        /// <remarks>
        /// If <see cref="Format"/> is set to <see cref="Ole10ObjectFormat.File"/> then this property is empty
        /// </remarks>
        public long Height { get; private set; }

        /// <summary>
        /// This MUST be a LengthPrefixedAnsiString or a LengthPrefixedUnicodeString, either of which contain a 
        /// registered clipboard format name
        /// </summary>
        /// <remarks>
        /// Only set when <see cref="Format"/> is set to <see cref="Ole10ObjectFormat.Presentation"/> and the
        /// <see cref="ClassName"/> does not contain "METAFILEPICT", "BITMAP" or "DIB"
        /// </remarks>
        public string StringFormatData { get; private set; }

        /// <summary>
        /// Identifies the <see cref="NativeData"/> when this file is a Clipboard object
        /// </summary>
        /// <remarks>
        /// Only set when <see cref="Format"/> is set to <see cref="Ole10ObjectFormat.Presentation"/> and the
        /// <see cref="ClassName"/> does not contain "METAFILEPICT", "BITMAP" or "DIB"
        /// </remarks>
        public ClipboardFormat ClipboardFormat { get; private set; }

        /// <summary>
        /// If the ObjectHeader structure is contained by an EmbeddedObject, 
        /// the TopicName field SHOULD contain an empty string and MUST be ignored on processing
        /// </summary>
        /// <remarks>
        /// If <see cref="Format"/> is set to <see cref="Ole10ObjectFormat.File"/> then this property is empty
        /// </remarks>
        public string TopicName { get; private set; }

        /// <summary>
        /// ItemName (variable): This MUST be a LengthPrefixedAnsiString.
        /// If the ObjectHeader structure is contained by an EmbeddedObject structure, 
        /// the ItemName field SHOULD contain an empty string and MUST be ignored on processing.
        /// If the ObjectHeader structure is contained by a LinkedObject structure,the ItemName field 
        /// MUST contain a string that is used by the application or higher-level protocol to identify 
        /// the item within the file to which is being linked. The format and meaning of the ItemName 
        /// string is specific to the creating application and MUST be treated by other parties as an 
        /// opaque string when processing this data structure. An example of such an item is an 
        /// individual cell within a spreadsheet application.
        /// </summary>
        /// <remarks>
        /// If <see cref="Format"/> is set to <see cref="Ole10ObjectFormat.File"/> then this property is empty
        /// </remarks>
        public string ItemName { get; private set; }

        /// <summary>
        /// If the <see cref="TopicName"/> field of the ObjectHeader structure contains a path that starts 
        /// with a drive letter and if the drive letter is for a remote drive, the NetworkName field MUST 
        /// contain the path name of the linked file in the Universal Naming Convention (UNC) format.
        /// </summary>
        /// <remarks>
        /// If <see cref="Format"/> is set to <see cref="Ole10ObjectFormat.File"/> then this property is empty
        /// </remarks>
        public string NetworkName { get; private set; }

        /// <summary>
        /// This field contains an implementation-specific hint supplied by the application or higher-level 
        /// protocol responsible for creating the data structure. The hint MAY be ignored on processing of 
        /// this data structure
        /// </summary>
        /// <remarks>
        /// If <see cref="Format"/> is set to <see cref="Ole10ObjectFormat.File"/> then this property is empty
        /// </remarks>
        public UInt32 LinkUpdateOptions { get; private set; }

        /// <summary>
        /// Data that is required to display the linked or embedded object within the container application.
        /// </summary>
        public byte[] PresentationData { get; private set; }

        /// <summary>
        /// The data that constitutes the state of an embedded object. The only entity that can create 
        /// and process the data is the creating application.
        /// </summary>
        public byte[] NativeData { get; private set; }
        #endregion

        #region ReadString
        /// <summary>
        /// Reads a length prefixd ansi or unicode string from the given <paramref name="binaryReader"/>
        /// <remarks>
        /// </remarks>
        /// Length (4 bytes): This MUST be set to the number of ANSI characters in the String field, 
        /// including the terminating null character. Length MUST be set to 0x00000000 to indicate an empty string.
        /// </summary>
        /// <param name="binaryReader"></param>
        /// <returns></returns>
        private static string ReadString(BinaryReader binaryReader)
        {
            var stringBuilder = new StringBuilder();
            var length = binaryReader.ReadUInt32();
            for(var i = 0; i< length; i++)
            {
                var b = binaryReader.ReadByte();
                if (b == 0)
                    return stringBuilder.ToString();

                // Unicode char found
                if (b >= 0xc2 && b <= 0xdf)
                {
                    var chr = new byte[2];
                    chr[1] = b;
                    var b2 = binaryReader.ReadByte();
                    chr[0] = b2;
                    stringBuilder.Append(Encoding.GetEncoding("ANSI6").GetString(chr));
                }
                else
                    stringBuilder.Append((char) b);
            }

            return stringBuilder.ToString();
        }
        #endregion

        #region Constructor
        /// <summary>
        /// Creates this object and sets all its properties
        /// </summary>
        /// <param name="inputStream">The OLE version 1.0 object as a stream</param>
        /// <exception cref="ArgumentNullException">Raised when <paramref name="inputStream"/> is <c>null</c></exception>
        public ObjectV10(Stream inputStream)
        {
            if (inputStream == null)
                throw new ArgumentNullException("inputStream");

            inputStream.Position = 0;

            using (var reader = new BinaryReader(inputStream))
                ParseOle10(reader);
        }

        /// <summary>
        /// Creates this object and sets all its properties
        /// </summary>
        /// <param name="reader">The OLE version 1.0 object as a stream</param>
        /// <exception cref="ArgumentNullException">Raised when <paramref name="reader"/> is <c>null</c></exception>
        internal ObjectV10(BinaryReader reader)
        {
            if (reader == null)
                throw new ArgumentNullException("reader");

            ParseOle10(reader);
        }        
        #endregion

        #region ParseOle10
        /// <summary>
        /// Parses the stream and sets all the OLE properties
        /// </summary>
        /// <param name="reader"></param>
        private void ParseOle10(BinaryReader reader)
        {
            Version = reader.ReadUInt32();

            var format = reader.ReadUInt32(); // FormatID
            try
            {
                Format = (Ole10ObjectFormat) format;
            }
            catch (Exception)
            {
                throw new OEFileIsCorrupt("Invalid OLE version 1.0 format, expected 0x00000000, 0x00000002 or 0x00000005");
            }

            if (Format != Ole10ObjectFormat.NotSet)
                ClassName = ReadString(reader);

            switch (Format)
            {
                case Ole10ObjectFormat.Link:
                    ParseObjectHeader(reader);
                    ParseLinkedObject(reader);
                    break;

                case Ole10ObjectFormat.File:
                    ParseObjectHeader(reader);
                    ParseEmbeddedObject(reader);
                    break;

                case Ole10ObjectFormat.Presentation:
                    switch (ClassName)
                    {
                        // MetaFilePresentationObject
                        case "METAFILEPICT":
                        case "BITMAP":
                        case "DIB":
                            ParseStandardPresentationObject(reader);
                            break;

                        default:
                            ParseGenericPresentationObject(reader);
                            break;
                    }

                    break;
        
                default:
                    throw new OEFileTypeNotSupported("Only OLE version 1.0 format 1 (linked object) and 2 (embedded object) files are supported");
            }
        }
        #endregion
       
        #region ParseStandardPresentationObject
        /// <summary>
        /// Parses the standard presentation object when the <see cref="Format"/> is set to <see cref="Ole10ObjectFormat.Presentation"/> and 
        /// the <see cref="ClassName"/> is set to "METAFILEPICT", "BITMAP" or "DIB"
        /// </summary>
        /// <param name="reader"></param>
        private void ParseStandardPresentationObject(BinaryReader reader)
        {
            Width = reader.ReadUInt32();
            Height = reader.ReadUInt32();

            switch (ClassName)
            {
                // MetaFilePresentationObject
                case "METAFILEPICT":
                {
                    var size = reader.ReadUInt32();

                    // PresentationDataSize (4 bytes): This MUST be an unsigned long integer set to the sum of the size,
                    // in bytes, of the PresentationData field and the number 8. If this field contains the value 8, 
                    // the PresentationData field MUST NOT be present.
                    if (size == 8)
                        return;

                    // Reserved1 (2 bytes): Reserved. This can be set to any arbitrary value and MUST be ignored on processing.
                    reader.ReadUInt16();
                    // Reserved2 (2 bytes): Reserved. This can be set to any arbitrary value and MUST be ignored on processing.
                    reader.ReadUInt16();
                    // Reserved3 (2 bytes): Reserved. This can be set to any arbitrary value and MUST be ignored on processing.
                    reader.ReadUInt16();
                    // Reserved4 (2 bytes): Reserved. This can be set to any arbitrary value and MUST be ignored on processing.
                    reader.ReadUInt16();

                    // This MUST be an array of bytes that contain a Windows metafile (as specified in [MS-WMF]).
                    if (size - 8 > 0)
                        PresentationData = reader.ReadBytes((int)size - 8);

                    break;
                }

                // BitmapPresentationObject
                case "BITMAP":
                case "DIB":
                {
                    var size = reader.ReadUInt32();

                    // PresentationDataSize (4 bytes): This MUST be an unsigned long integer set to the size, 
                    // in bytes, of the Bitmap or DIB field. If this field has the value 0, the Bitmap or DIB field MUST 
                    // NOT be present.
                    if (size == 0)
                        return;

                    PresentationData = reader.ReadBytes((int)size);
                    break;
                }

                default:
                    throw new OEFileTypeNotSupported(
                        "Unexpected value found in classname field, expected METAFILEPICT, BITMAP or DIB");
            }            
        }
        #endregion

        #region ParseGenericPresentationObject
        /// <summary>
        /// Parses the generic presentation object when the <see cref="Format"/> is set to <see cref="Ole10ObjectFormat.Presentation"/> and 
        /// the <see cref="ClassName"/> is <b>NOT</b> set to "METAFILEPICT", "BITMAP" or "DIB"
        /// </summary>
        /// <param name="reader"></param>
        private void ParseGenericPresentationObject(BinaryReader reader)
        {
            ClipboardFormat = (ClipboardFormat) reader.ReadUInt32();
            
            switch (ClipboardFormat)
            {
                // RegisteredClipboardFormatPresentationObject
                case ClipboardFormat.Registered:
                {
                    // This MUST be set to the size, in bytes, of the StringFormatData field.
                    // ReSharper disable once UnusedVariable
                    var stringFormatDataSize = reader.ReadUInt32();
                    StringFormatData = ReadString(reader);
                    var size = reader.ReadUInt32();
                    NativeData = reader.ReadBytes((int)size);
                    break;
                }

                case ClipboardFormat.CF_BITMAP:
                case ClipboardFormat.CF_DIB:
                case ClipboardFormat.CF_ENHMETAFILE:
                case ClipboardFormat.CF_METAFILEPICT:
                {
                    var size = reader.ReadUInt32();
                    NativeData = reader.ReadBytes((int) size);
                    break;
                }

                default:
                    throw new OEFileTypeNotSupported(
                        "A not supported clipboardformat has been found, only CF_BITMAP, CF_DIB, CF_ENHMETAFILE and CF+METAFILEPICT are supported");
            }            
        }
        #endregion

        #region ParseObjectHeader
        /// <summary>
        /// Parses the stream when the <see cref="Format"/> is set to <see cref="Ole10ObjectFormat.File"/>
        /// or set to <see cref="Ole10ObjectFormat.Link"/>
        /// </summary>
        /// <param name="reader"></param>
        private void ParseObjectHeader(BinaryReader reader)
        {
            TopicName = ReadString(reader);
            ItemName = ReadString(reader);
        }
        #endregion

        #region GetAndSetPresentationObject
        /// <summary>
        /// A <see cref="ParseEmbeddedObject"/> or <see cref="ParseLinkedObject"/> always contain
        /// a presentation object, this is used to present the embedded link or object in the host
        /// application. This method will extract this data.
        /// </summary>
        /// <param name="reader"></param>
        public void GetAndSetPresentationObject(BinaryReader reader)
        {
            var po = new ObjectV10(reader);
            Width = po.Width;
            Height = po.Height;
            StringFormatData = po.StringFormatData;
            ClipboardFormat = po.ClipboardFormat;
            PresentationData = po.PresentationData;            
        }
        #endregion

        #region ParseEmbeddedObject
        /// <summary>
        /// Parses the stream when it is an <see cref="Ole10ObjectFormat.File"/>
        /// </summary>
        /// <param name="reader"></param>
        private void ParseEmbeddedObject(BinaryReader reader)
        {
            var nativeDataSize = reader.ReadUInt32();
            NativeData = reader.ReadBytes((int) nativeDataSize);
            GetAndSetPresentationObject(reader);
        }
        #endregion

        #region ParseLinkedObject
        /// <summary>
        /// Parses the stream when it is an <see cref="Ole10ObjectFormat.Link"/>
        /// </summary>
        /// <param name="reader"></param>
        private void ParseLinkedObject(BinaryReader reader)
        {
            NetworkName = ReadString(reader);
            TopicName = ReadString(reader);
            // Reserved (4 bytes)
            reader.ReadUInt32();
            LinkUpdateOptions = reader.ReadUInt32();
            GetAndSetPresentationObject(reader);
        }
        #endregion
    }
}