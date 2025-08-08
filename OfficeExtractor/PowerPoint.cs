using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;
using OpenMcdf;

//
// PowerPoint.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2025 Magic-Sessions. (www.magic-sessions.com)
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

namespace OfficeExtractor
{
    /// <summary>
    /// This class is used as a placeholder for all PowerPoint related methods
    /// </summary>
    internal class PowerPoint
    {
        #region Fields
        /// <summary>
        ///     <see cref="Extraction"/>
        /// </summary>
        private Extraction _extraction;
        #endregion

        #region Properties
        /// <summary>
        /// Returns a reference to the Extraction class when it already exists or creates a new one
        /// when it doesn't
        /// </summary>
        private Extraction Extraction
        {
            get
            {
                if (_extraction != null)
                    return _extraction;

                _extraction = new Extraction();
                return _extraction;
            }
        }
        #endregion

        #region Extract
        /// <summary>
        /// This method saves all the PowerPoint embedded binary objects from the <paramref name="inputFile"/> to the
        /// <paramref name="outputFolder"/>
        /// </summary>
        /// <param name="inputFile">The binary PowerPoint file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns></returns>
        /// <exception cref="OEFileIsPasswordProtected">Raised when the <paramref name="inputFile"/> is password protected</exception>
        internal List<string> Extract(string inputFile, string outputFolder)
        {
            Logger.WriteToLog("The file is a binary PowerPoint document");

            using var compoundFile = RootStorage.OpenRead(inputFile);
            var result = new List<string>();
            if (!compoundFile.TryOpenStream("PowerPoint Document", out var stream))
            {
                Logger.WriteToLog("Could not find the PowerPoint Document stream inside the file");
                return result;
            }

            Logger.WriteToLog("PowerPoint Document stream found");

            using var binaryReader = new BinaryReader(stream);
            while (binaryReader.BaseStream.Position != stream.Length)
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
                        {
                            var fileName = outputFolder + Extraction.DefaultEmbeddedObjectName;
                            var extractedFileName = Extraction.SaveByteArrayToFile(bytes, fileName);

                            if (!string.IsNullOrEmpty(extractedFileName))
                                result.Add(extractedFileName);
                        }
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
                        Logger.WriteToLog("Uncompressing byte array");
                        var deflateStream = new DeflateStream(compressedMemoryStream, CompressionMode.Decompress, true);

                        // NET6 had a breaking change in DeflateStream (et.al.): calling int Read(..) tends to
                        // not read until the desired count but only fewer bytes, one must rely on returned count
                        var decrByteCount = 0;
                        while (decrByteCount < decompressedSize)
                            decrByteCount += deflateStream.Read(decompressedBytes, decrByteCount, (int)(decompressedSize - decrByteCount));

                        Logger.WriteToLog("Byte array uncompressed");

                        string extractedFileName;

                        // Check if the ole object is another compound storage node with a package stream
                        if (Extraction.IsCompoundFile(decompressedBytes))
                            extractedFileName = Extraction.SaveFromStorageNode(decompressedBytes, outputFolder);
                        else
                        {
                            var fileName = outputFolder + Extraction.DefaultEmbeddedObjectName;
                            extractedFileName = Extraction.SaveByteArrayToFile(decompressedBytes, fileName);
                        }

                        if (!string.IsNullOrEmpty(extractedFileName))
                            result.Add(extractedFileName);
                    }
                }
                else
                    binaryReader.BaseStream.Position += size;
            }

            return result;
        }
        #endregion
    }
}
