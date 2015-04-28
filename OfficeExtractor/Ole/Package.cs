using System.IO;
using CompoundFileStorage;
using CompoundFileStorage.Interfaces;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;

namespace OfficeExtractor.Ole
{
    internal class Package
    {
        #region Properties
        /// <summary>
        ///     The name of the file
        /// </summary>
        public string FileName { get; private set; }

        /// <summary>
        ///     The original location of the file (before it was embedded)
        /// </summary>
        public string FilePath { get; private set; }

        /// <summary>
        ///     The original location of the file (before it was embedded)
        /// </summary>
        public string TemporaryPath { get; private set; }
        
        /// <summary>
        /// The file data
        /// </summary>
        public byte[] Data { get; private set; }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets all its properties
        /// </summary>
        /// <param name="stream">The Compound File Storage CompObj <see cref="CFStream" /></param>
        internal Package(ICFStream stream)
        {
            using (var memoryStream = new MemoryStream(stream.GetData()))
            using (var binaryReader = new BinaryReader(memoryStream))
            {
                // Skip the first byte
                binaryReader.ReadByte();

                // Check signature
                var signature = binaryReader.ReadUInt16();
                if (signature != 0x0002)
                    throw new OEFileIsCorrupt("Invalid package type signature, expected 0x0002");

                if (binaryReader.PeekChar() == 00)
                    binaryReader.ReadByte();

                // Check if we have a double signature. In this case the FileName and FilePath are
                // also added at the end of the file in unicode format
                var signatureUnicode = binaryReader.ReadUInt16();

                FileName = Strings.ReadNullTerminatedAnsiString(binaryReader);
                FilePath = Strings.ReadNullTerminatedAnsiString(binaryReader);

                // Skip 2 unused bytes
                binaryReader.ReadBytes(2);

                var temp = binaryReader.ReadUInt16();
                if (temp == 0x0003)
                    TemporaryPath = Strings.Read4ByteLengthPrefixedAnsiString(binaryReader);

                var dataSize = (int) binaryReader.ReadUInt32();

                // And finaly we have come to the original file
                Data = binaryReader.ReadBytes(dataSize);

                // If a double signature was found then also read the unicode parts at the
                // end of the file
                if (signatureUnicode == 0x0002)
                {
                    FileName = Strings.Read4ByteLengthPrefixedUnicodeString(binaryReader);
                    FilePath = Strings.Read4ByteLengthPrefixedUnicodeString(binaryReader);
                    TemporaryPath = Strings.Read4ByteLengthPrefixedUnicodeString(binaryReader);
                }
            }
        }
        #endregion
    }
}