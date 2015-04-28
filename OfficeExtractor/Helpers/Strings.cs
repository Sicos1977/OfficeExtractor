using System.IO;
using System.Text;

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

            while (binaryReader.PeekChar() != -1)
            {
                var b = binaryReader.ReadByte();
                if (b == 0)
                    return stringBuilder.ToString();

                stringBuilder.Append((char)b);
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