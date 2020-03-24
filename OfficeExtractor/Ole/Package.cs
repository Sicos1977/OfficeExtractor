using System.IO;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;

//
// Package.cs
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
                var tempFileName = Strings.Read4ByteLengthPrefixedUnicodeString(binaryReader);
                if (string.IsNullOrEmpty(FileName)) FileName = tempFileName;
                var tempFilePath = Strings.Read4ByteLengthPrefixedUnicodeString(binaryReader);
                if (string.IsNullOrEmpty(FilePath)) FilePath = tempFilePath;
                TemporaryPath = Strings.Read4ByteLengthPrefixedUnicodeString(binaryReader);
            }            
        }
        #endregion
    }
}