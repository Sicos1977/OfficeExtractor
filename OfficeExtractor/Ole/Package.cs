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
        ///     This MUST be set to <see cref="OleFormat.Link" /> (0x00000001) or <see cref="OleFormat.File" />
        ///     (0x00000002).
        /// </summary>
        public OleFormat Format { get; private set; }

        /// <summary>
        ///     When <see cref="Format"/> is set to <see cref="OleFormat.File"/> then this will contain the original
        ///     name of the embedded file. When set to <see cref="OleFormat.Link"/> this wil contain the name of the
        ///     linked file.
        /// </summary>
        public string FileName { get; private set; }

        /// <summary>
        ///     When <see cref="Format"/> is set to <see cref="OleFormat.File"/> then this will contain the original
        ///     location of the embedded file. When set to <see cref="OleFormat.Link"/> this wil contain the path to
        ///     the linked file.
        /// </summary>
        public string FilePath { get; private set; }

        /// <summary>
        ///     When <see cref="Format"/> is set to <see cref="OleFormat.File"/> then this will contain the temporary
        ///     location that was used to embedded the file. When set to <see cref="OleFormat.Link"/> this wil contain 
        ///     the path to the linked file (the same as <see cref="FilePath"/>).
        /// </summary>
        public string TemporaryPath { get; private set; }
        
        /// <summary>
        ///     The file data
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
                // Skip the first 4 bytes, this contains the ole10Native data length size
                binaryReader.ReadUInt32();

                // Check signature
                var signature = binaryReader.ReadUInt16();
                if (signature != 0x0002)
                    throw new OEFileIsCorrupt("Invalid package type signature, expected 0x0002");

                if (binaryReader.PeekChar() == 00)
                    binaryReader.ReadByte();

                FileName = Path.GetFileName(Strings.ReadNullTerminatedAnsiString(binaryReader));
                FilePath = Strings.ReadNullTerminatedAnsiString(binaryReader);

                // Skip 2 unused bytes
                binaryReader.ReadBytes(2);

                var format = binaryReader.ReadUInt16();

                switch (format)
                {
                    case 0x00000001:
                        Format = OleFormat.Link;
                        break;

                    case 0x00000003:
                        Format = OleFormat.File;
                        var dataSize = (int) binaryReader.ReadUInt32();
                        Data = binaryReader.ReadBytes(dataSize);
                        break;

                    default:
                        throw new OEObjectTypeNotSupported("Invalid signature found, expected 0x00000001 or 0x00000003");
                }

                TemporaryPath = Strings.Read4ByteLengthPrefixedAnsiString(binaryReader);

                if (binaryReader.BaseStream.Position >= binaryReader.BaseStream.Length) return;
                FileName = Strings.Read4ByteLengthPrefixedUnicodeString(binaryReader);
                FilePath = Strings.Read4ByteLengthPrefixedUnicodeString(binaryReader);
                TemporaryPath = Strings.Read4ByteLengthPrefixedUnicodeString(binaryReader);
            }
        }
        #endregion
    }
}