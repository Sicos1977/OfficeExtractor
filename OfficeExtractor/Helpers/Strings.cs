using System;
using System.IO;
using System.Text;

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
            var stringBuilder = new StringBuilder();

            while (binaryReader.BaseStream.Position != 
                   binaryReader.BaseStream.Length)
            {
                var b = binaryReader.ReadByte();
                if (b == 0)
                    return stringBuilder.ToString();

                stringBuilder.Append(Convert.ToChar(b));
            }

            return stringBuilder.ToString();
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