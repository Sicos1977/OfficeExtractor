using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using CompoundFileStorage;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;

/*
   Copyright 2014-2016 Kees van Spelde

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

namespace OfficeExtractor
{
    /// <summary>
    /// This class is used as a placeholder for all PowerPoint related methods
    /// </summary>
    internal static class PowerPoint
    {
        #region SaveToFolder
        /// <summary>
        /// This method saves all the PowerPoint embedded binary objects from the <paramref name="inputFile"/> to the
        /// <paramref name="outputFolder"/>
        /// </summary>
        /// <param name="inputFile">The binary PowerPoint file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns></returns>
        /// <exception cref="OEFileIsPasswordProtected">Raised when the <see cref="inputFile"/> is password protected</exception>
        public static List<string> SaveToFolder(string inputFile, string outputFolder)
        {
            using (var compoundFile = new CompoundFile(inputFile))
            {
                if (IsPasswordProtected(compoundFile))
                    throw new OEFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                        "' is password protected");

                var result = new List<string>();

                if (!compoundFile.RootStorage.ExistsStream("PowerPoint Document")) return result;
                var stream = compoundFile.RootStorage.GetStream("PowerPoint Document") as CFStream;
                if (stream == null) return result;

                using (var memoryStream = new MemoryStream(stream.GetData()))
                using (var binaryReader = new BinaryReader(memoryStream))
                {
                    while (binaryReader.BaseStream.Position != memoryStream.Length)
                    {
                        var verAndInstance = binaryReader.ReadUInt16();
                        // ReSharper disable once UnusedVariable
                        var version = verAndInstance & 0x000FU; // First 4 bit of field verAndInstance
                        var instance = (verAndInstance & 0xFFF0U) >> 4; // Last 12 bit of field verAndInstance

                        var typeCode = binaryReader.ReadUInt16();
                        var size = binaryReader.ReadUInt32();

                        // Embedded OLE objects start with code 4113
                        if (typeCode == 4113)
                        {
                            if (instance == 0)
                            {
                                // Uncompressed
                                var bytes = binaryReader.ReadBytes((int)size);

                                // Check if the ole object is another compound storage node with a package stream
                                if (Extraction.IsCompoundFile(bytes))
                                    result.Add(Extraction.SaveFromStorageNode(bytes, outputFolder));
                                else
                                    Extraction.SaveByteArrayToFile(bytes, outputFolder + Extraction.DefaultEmbeddedObjectName);
                            }
                            else
                            {
                                var decompressedSize = binaryReader.ReadUInt32();
                                var data = binaryReader.ReadBytes((int)size - 4);
                                var compressedMemoryStream = new MemoryStream(data);

                                // skip the first 2 bytes
                                compressedMemoryStream.ReadByte();
                                compressedMemoryStream.ReadByte();

                                // Decompress the bytes
                                var decompressedBytes = new byte[decompressedSize];
                                var deflateStream = new DeflateStream(compressedMemoryStream, CompressionMode.Decompress, true);
                                deflateStream.Read(decompressedBytes, 0, decompressedBytes.Length);

                                string extractedFileName;

                                // Check if the ole object is another compound storage node with a package stream
                                if (Extraction.IsCompoundFile(decompressedBytes))
                                    extractedFileName = Extraction.SaveFromStorageNode(decompressedBytes, outputFolder);
                                else
                                    extractedFileName = Extraction.SaveByteArrayToFile(decompressedBytes,
                                        outputFolder + Extraction.DefaultEmbeddedObjectName);

                                if (!string.IsNullOrEmpty(extractedFileName))
                                    result.Add(extractedFileName);
                            }
                        }
                        else
                            binaryReader.BaseStream.Position += size;
                    }
                }

                return result;
            }
        }
        #endregion

        #region IsPasswordProtected
        /// <summary>
        /// Returns true when the binary PowerPoint file is password protected
        /// </summary>
        /// <param name="compoundFile"></param>
        /// <returns></returns>
        private static bool IsPasswordProtected(CompoundFile compoundFile)
        {
            if (!compoundFile.RootStorage.ExistsStream("Current User")) return false;
            var stream = compoundFile.RootStorage.GetStream("Current User") as CFStream;
            if (stream == null) return false;

            using (var memoryStream = new MemoryStream(stream.GetData()))
            using (var binaryReader = new BinaryReader(memoryStream))
            {
                var verAndInstance = binaryReader.ReadUInt16();
                // ReSharper disable UnusedVariable
                // We need to read these fields to get to the correct location in the Current User stream
                var version = verAndInstance & 0x000FU;         // first 4 bit of field verAndInstance
                var instance = (verAndInstance & 0xFFF0U) >> 4; // last 12 bit of field verAndInstance
                var typeCode = binaryReader.ReadUInt16();
                var size = binaryReader.ReadUInt32();
                var size1 = binaryReader.ReadUInt32();
                // ReSharper restore UnusedVariable
                var headerToken = binaryReader.ReadUInt32();

                switch (headerToken)
                {
                    // Not encrypted
                    case 0xE391C05F:
                        return false;

                    // Encrypted
                    case 0xF3D1C4DF:
                        return true;

                    default:
                        return false;
                }
            }
        }
        #endregion
    }
}
