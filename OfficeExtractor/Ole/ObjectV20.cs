using System;
using System.IO;
using System.Text;

namespace OfficeExtractor.Ole
{
    /// <summary>
    /// This class represents an OLE version 2.0 object
    /// </summary>
    /// <remarks>
    /// See the Microsoft documentation at https://msdn.microsoft.com/en-us/library/dd942280.aspx
    /// </remarks>
    internal class ObjectV20
    {
        #region Properties
        /// <summary>
        /// The signature of the file
        /// </summary>
        public UInt16 Signature { get; private set; }

        /// <summary>
        /// The name of the file
        /// </summary>
        public string FileName { get; private set; }

        /// <summary>
        /// The original locatation of the file (before it was embedded)
        /// </summary>
        public string FilePath { get; private set; }

        /// <summary>
        /// The file data
        /// </summary>
        public byte[] Data { get; private set; }
        #endregion

        #region ReadString
        /// <summary>
        /// Reads an fixed length ansi string from the given <paramref name="inputStream"/>
        /// </summary>
        /// <param name="inputStream"></param>
        /// <returns></returns>
        private static string ReadString(Stream inputStream)
        {
            var result = string.Empty;

            int i;
            var chr = new byte[1];
            for (i = 0; i < 260; i++)
            {
                inputStream.Read(chr, 0, 1);
                if (chr[0] == 0)
                    break;

                // Unicode char found
                if (chr[0] >= 0xc2 && chr[0] <= 0xdf)
                {
                    i += 1;

                    var chr2 = new byte[2];
                    chr2[1] = chr[0];
                    inputStream.Read(chr, 0, 1);
                    chr2[0] = chr[0];

                    result += Encoding.GetEncoding("ANSI6").GetString(chr2);
                }
                else
                    result += (char)chr[0];
            }

            return result;
        }
        #endregion

        #region Constructor
        /// <summary>
        /// Creates this object and sets all its properties
        /// </summary>
        /// <param name="inputStream">The OLE version 2.0 object as a stream</param>
        public ObjectV20(Stream inputStream)
        {
            if (inputStream == null)
                throw new ArgumentNullException("inputStream");

            inputStream.Position = 0;

            var binaryReader = new BinaryReader(inputStream);
            Signature = binaryReader.ReadUInt16(); // Signature

            // The name of the file start at postion 7 so move to there
            inputStream.Position += 4;

            FileName = ReadString(inputStream);
            FilePath = ReadString(inputStream);

            // We need to skip the next four bytes
            inputStream.Position += 4;

            // Read the tempory path size
            var size = new byte[4];
            inputStream.Read(size, 0, 4);
            var tempPathSize = BitConverter.ToInt32(size, 0);

            // Move the position in the stream after the temp path
            inputStream.Position += tempPathSize;

            // Read the next four bytes for the length of the data
            inputStream.Read(size, 0, 4);
            var fileSize = BitConverter.ToInt32(size, 0);

            // And finaly we have come to the original file
            Data = new byte[fileSize];
            inputStream.Read(Data, 0, fileSize);
        }
        #endregion
    }
}