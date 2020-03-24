using System.IO;
using System.Text;

//
// MonikerStream.cs
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
    /// <summary>
    ///     The MONIKERSTREAM structure specifies the reference to the linked object.
    /// </summary>
    internal class MonikerStream
    {
        #region Properties
        /// <summary>
        ///     This MUST be the packetized CLSID of an implementation-specific object capable of processing the
        ///     data contained in the StreamData field.
        /// </summary>
        public CLSID Clsid { get; private set; }

        /// <summary>
        ///     This MUST be an array of bytes that specifies the reference to the linked object. The value of
        ///     this array is interpreted in an implementation-specific manner.
        /// </summary>
        public byte[] StreamData { get; private set; }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets all its properties
        /// </summary>
        /// <param name="binaryReader"></param>
        /// <param name="size">The size of the monikerstream</param>
        internal MonikerStream(BinaryReader binaryReader, int size)
        {
            // 16 bytes
            Clsid = new CLSID(binaryReader);
            StreamData = binaryReader.ReadBytes(size - 16);
        }
        #endregion

        #region ToString
        /// <summary>
        ///     Returns the <see cref="StreamData" /> as a string
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Encoding.UTF8.GetString(StreamData);
        }
        #endregion
    }
}