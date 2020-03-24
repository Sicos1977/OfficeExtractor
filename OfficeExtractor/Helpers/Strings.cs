using System.IO;
using System.Text;

//
// Strings.cs
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

namespace OfficeExtractor.Helpers
{
    internal static class Strings
    {
        #region ReadNullTerminatedAnsiString
        /// <summary>
        ///     Reads an null terminated ansi string from given <paramref name="binaryReader"/> until the 
        ///     null char has been found
        /// </summary>
        /// <param name="binaryReader"></param>
        /// <returns></returns>
        internal static string ReadNullTerminatedAnsiString(BinaryReader binaryReader)
        {
            var bytes = new byte[1024];
            var index = 0;

            do
            {
                var b = binaryReader.ReadByte();
                bytes[index++] = b;
                if (b == 0)
                    break;

            } while (true);

            var bytesUnicode = Encoding.Convert(Encoding.Default, Encoding.Unicode, bytes);
            var str = Encoding.Unicode.GetString(bytesUnicode);
            return str.TrimEnd('\0');
        }
        #endregion

        #region Read1ByteLengthPrefixedAnsiString
        /// <summary>
        ///     Reads a 1 byte length prefixed ansi string from the given <paramref name="binaryReader" />
        /// </summary>
        /// <param name="binaryReader"></param>
        /// <returns></returns>
        internal static string Read1ByteLengthPrefixedAnsiString(BinaryReader binaryReader)
        {
            var length = (int) binaryReader.ReadByte();
            var bytes = binaryReader.ReadBytes(length);
            var str = Encoding.UTF8.GetString(bytes);
            return str.TrimEnd('\0');
        }
        #endregion

        #region Read4ByteLengthPrefixedAnsiString
        /// <summary>
        ///     Reads a 4 byte length prefixed ansi string from the given <paramref name="binaryReader" />
        /// </summary>
        /// <param name="binaryReader"></param>
        /// <returns></returns>
        internal static string Read4ByteLengthPrefixedAnsiString(BinaryReader binaryReader)
        {
            var length = (int) binaryReader.ReadUInt32();
            var bytes = binaryReader.ReadBytes(length);
            var str = Encoding.UTF8.GetString(bytes);
            return str.TrimEnd('\0');
        }
        #endregion

        #region Read4ByteLengthPrefixedUnicodeString
        /// <summary>
        ///     Reads a 4 byte length prefixed ansi string from the given <paramref name="binaryReader" />
        /// </summary>
        /// <param name="binaryReader"></param>
        /// <returns></returns>
        internal static string Read4ByteLengthPrefixedUnicodeString(BinaryReader binaryReader)
        {
            var length = (int)binaryReader.ReadUInt32() * 2;
            var bytes = binaryReader.ReadBytes(length);
            var str = Encoding.Unicode.GetString(bytes);
            return str.TrimEnd('\0');
        }
        #endregion
    }
}