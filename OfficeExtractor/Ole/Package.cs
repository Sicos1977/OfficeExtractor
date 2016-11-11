using System.IO;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;

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
        /// <param name="data">The Package object as an byte array</param>
        internal Package(byte[] data)
        {
            ParsePackage(data);
        }
        #endregion

        #region ParsePackage
        /// <summary>
        ///     Parses the byte array and sets all the package properties
        /// </summary>
        /// <param name="data"></param>
        private void ParsePackage(byte[] data)
        {
            using (var memoryStream = new MemoryStream(data))
            using (var binaryReader = new BinaryReader(memoryStream))
            {
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

                // Read format
                var format = binaryReader.ReadUInt16();

                // Read temporary path
                TemporaryPath = Strings.Read4ByteLengthPrefixedAnsiString(binaryReader);

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
                
                if (binaryReader.BaseStream.Position >= binaryReader.BaseStream.Length) return;
                FileName = Strings.Read4ByteLengthPrefixedUnicodeString(binaryReader);
                FilePath = Strings.Read4ByteLengthPrefixedUnicodeString(binaryReader);
                TemporaryPath = Strings.Read4ByteLengthPrefixedUnicodeString(binaryReader);
            }            
        }
        #endregion
    }
}