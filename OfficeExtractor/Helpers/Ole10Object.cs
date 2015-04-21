using System;
using System.IO;
using System.Text;

namespace OfficeExtractor.Helpers
{
    #region Ole10ObjectType
    /// <summary>
    /// Type OLE version 1.0 object type
    /// </summary>
    internal enum Ole10ObjectType
    {
        /// <summary>
        /// The type is unknown
        /// </summary>
        Unknown = 0,

        /// <summary>
        /// The packaged object is a link
        /// </summary>
        Link = 1,

        /// <summary>
        /// The packaged object is a file
        /// </summary>
        File = 3
    }
    #endregion

    /// <summary>
    /// This class represents an OLE version 1.0 object
    /// <remarks>
    /// See the Microsoft documentation at https://msdn.microsoft.com/en-us/library/dd942402.aspx
    /// </remarks>
    /// </summary>
    internal class Ole10Object
    {
        #region Properties
        /// <summary>
        /// The signature of the file
        /// </summary>
        public UInt16 Signature { get; private set; }

        /// <summary>
        /// The displayname for the object
        /// </summary>
        public string DisplayName { get; private set; }
        
        /// <summary>
        /// The path to the icon for the file that is inside the packaged object
        /// </summary>
        public string IconFilePath { get; private set; }

        /// <summary>
        /// The position index for the icon in the RTF file
        /// </summary>
        public int IconIndex { get; private set; }

        /// <summary>
        /// The name of the packaged file
        /// </summary>
        public string FilePath { get; private set; }

        /// <summary>
        /// The type
        /// </summary>
        public Ole10ObjectType Type { get; private set; }

        /// <summary>
        /// The file data
        /// </summary>
        public byte[] Data { get; private set; }
        #endregion

        #region ReadAnsiString
        /// <summary>
        /// Reads an ansi string from the given <paramref name="binaryReader"/>
        /// </summary>
        /// <param name="binaryReader"></param>
        /// <returns></returns>
        private static string ReadAnsiString(BinaryReader binaryReader)
        {
            var stringBuilder = new StringBuilder();
            do
            {
                var b = binaryReader.ReadByte();
                if (b == 0)
                    return stringBuilder.ToString();

                stringBuilder.Append((char) b);
            } while (true);
        }
        #endregion

        #region Constructor
        /// <summary>
        /// Creates this object and sets all its properties
        /// </summary>
        /// <param name="inputStream">The OLE version 1.0 object as a stream</param>
        public Ole10Object(Stream inputStream)
        {
            if (inputStream == null)
                throw new ArgumentNullException("inputStream");

            var reader = new BinaryReader(inputStream);
            Signature = reader.ReadUInt16(); // Signature
            DisplayName = ReadAnsiString(reader);
            IconFilePath = ReadAnsiString(reader);
            IconIndex = reader.ReadUInt16();

            var type = reader.ReadUInt16();
            if (type == 1 || type == 3)
                Type = (Ole10ObjectType) type;
            else
                Type = Ole10ObjectType.Unknown;

            reader.ReadInt32(); // Nextsize
            FilePath = ReadAnsiString(reader);
            var dataSize = reader.ReadInt32();
            Data = reader.ReadBytes(dataSize);
        }
        #endregion
    }
}