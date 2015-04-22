using System;
using System.IO;
using System.Text;
using OfficeExtractor.Exceptions;

namespace OfficeExtractor.Ole
{
    #region Ole10ObjectFormat
    /// <summary>
    /// Type OLE version 1.0 object type
    /// </summary>
    internal enum Ole10ObjectFormat
    {
        /// <summary>
        /// The format is unknown
        /// </summary>
        Unknown = 0,

        /// <summary>
        /// The embedded object is a link
        /// </summary>
        Link = 1,

        /// <summary>
        /// The embedded object is a file
        /// </summary>
        File = 2
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
        /// This MUST be set to 0x00000001 or 0x00000002. Otherwise, the ObjectHeader structure is invalid
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
        /// If the ObjectHeader structure is contained by an EmbeddedObject, 
        /// the TopicName field SHOULD contain an empty string and MUST be ignored on processing
        /// </summary>
        /// <remarks>
        /// If <see cref="Format"/> is set to 0x00000002 then this property is empty
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
        /// If <see cref="Format"/> is set to 0x00000002 then this property is empty
        /// </remarks>
        public string ItemName { get; private set; }

        /// <summary>
        /// If the <see cref="TopicName"/> field of the ObjectHeader structure contains a path that starts 
        /// with a drive letter and if the drive letter is for a remote drive, the NetworkName field MUST 
        /// contain the path name of the linked file in the Universal Naming Convention (UNC) format.
        /// </summary>
        public string NetworkName { get; private set; }

        /// <summary>
        /// The position index for the icon in the RTF file
        /// </summary>
        public int IconIndex { get; private set; }

        /// <summary>
        /// The original locatation of the file (before it was embedded)
        /// </summary>
        public string FilePath { get; private set; }

        /// <summary>
        /// The file data
        /// </summary>
        public byte[] Data { get; private set; }
        #endregion

        #region ReadAnsiString
        /// <summary>
        /// Reads a length prefixd ansi string from the given <paramref name="binaryReader"/>
        /// <remarks>
        /// </remarks>
        /// Length (4 bytes): This MUST be set to the number of ANSI characters in the String field, 
        /// including the terminating null character. Length MUST be set to 0x00000000 to indicate an empty string.
        /// </summary>
        /// <param name="binaryReader"></param>
        /// <returns></returns>
        private static string ReadAnsiString(BinaryReader binaryReader)
        {
            var stringBuilder = new StringBuilder();
            var length = binaryReader.ReadUInt32();
            for(var i = 0; i< length; i++)
            {
                var b = binaryReader.ReadByte();
                if (b == 0)
                    return stringBuilder.ToString();

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
        public ObjectV10(Stream inputStream)
        {
            if (inputStream == null)
                throw new ArgumentNullException("inputStream");

            inputStream.Position = 0;

            var reader = new BinaryReader(inputStream);
            Version = reader.ReadUInt32();
            // FormatID (4 bytes): This MUST be set to 0x00000001 or 0x00000002. Otherwise, the ObjectHeader structure is invalid
            var format = reader.ReadUInt32(); // FormatID
            if (format == 1 || format == 2)
                Format = (Ole10ObjectFormat) format;
            else
                Format = Ole10ObjectFormat.Unknown;

            ClassName = ReadAnsiString(reader);
            TopicName = ReadAnsiString(reader);
            ItemName = ReadAnsiString(reader);

            switch (Format)
            {
                case Ole10ObjectFormat.Link:
                    NetworkName = ReadAnsiString(reader);
                    break;

                case Ole10ObjectFormat.File:
                    break;

                default:
                    throw new OEFileTypeNotSupported("Only OLE version 1.0 format 1 (linked object) and 2 (embedded object) files are supported");
            }

            IconIndex = reader.ReadUInt16();


            reader.ReadInt32(); // Nextsize
            FilePath = ReadAnsiString(reader);
            var dataSize = reader.ReadInt32();
            Data = reader.ReadBytes(dataSize);
        }
        #endregion
    }
}