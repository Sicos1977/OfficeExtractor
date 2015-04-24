using System.IO;
using System.Text;

namespace OfficeExtractor.Helpers
{
    internal static class Strings
    {
        #region Read1ByteLengthPrefixedString
        /// <summary>
        /// Reads a 1 byte length prefixd ansi string from the given <paramref name="binaryReader"/>
        /// <remarks>
        /// </remarks>
        /// Length (4 bytes): This MUST be set to the number of ANSI characters in the String field, 
        /// including the terminating null character. Length MUST be set to 0x00000000 to indicate an empty string.
        /// </summary>
        /// <param name="binaryReader"></param>
        /// <returns></returns>
        internal static string Read1ByteLengthPrefixedString(BinaryReader binaryReader)
        {
            var stringBuilder = new StringBuilder();
            var length = binaryReader.ReadByte();

            for (var i = 0; i < length; i++)
            {
                var b = binaryReader.ReadByte();
                if (b == 0)
                    return stringBuilder.ToString();

                stringBuilder.Append((char)b);
            }

            return stringBuilder.ToString();
        }
        #endregion

        #region Read4ByteLengthPrefixedString
        /// <summary>
        /// Reads a 2 byte length prefixd ansi or unicode string from the given <paramref name="binaryReader"/>
        /// <remarks>
        /// </remarks>
        /// Length (4 bytes): This MUST be set to the number of ANSI characters in the String field, 
        /// including the terminating null character. Length MUST be set to 0x00000000 to indicate an empty string.
        /// </summary>
        /// <param name="binaryReader"></param>
        /// <returns></returns>
        internal static string Read4ByteLengthPrefixedString(BinaryReader binaryReader)
        {
            var stringBuilder = new StringBuilder();
            var length = binaryReader.ReadUInt32();
            
            for (var i = 0; i < length; i++)
            {
                var b = binaryReader.ReadByte();
                if (b == 0)
                    return stringBuilder.ToString();

                // Unicode char found
                if (b >= 0xc2 && b <= 0xdf)
                {
                    var chr = new byte[2];
                    chr[1] = b;
                    var b2 = binaryReader.ReadByte();
                    chr[0] = b2;
                    stringBuilder.Append(Encoding.GetEncoding(1255).GetString(chr));
                }
                else
                    stringBuilder.Append((char)b);
            }

            return stringBuilder.ToString();
        }
        #endregion
    }
}
