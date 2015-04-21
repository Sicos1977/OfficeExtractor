using System;
using System.IO;
using OfficeExtractor.Biff8.Interfaces;
using OfficeExtractor.Exceptions;

namespace OfficeExtractor.Biff8
{
    /// <summary>
    ///     Wraps an <see cref="T:System.IO.Stream" /> providing <see cref="T:NPOI.Util.ILittleEndianInput" /><p />
    ///     This class does not buffer any input, so the stream Read position maintained
    ///     by this class is consistent with that of the inner stream.
    /// </summary>
    /// <remarks>
    ///     @author Josh Micich
    /// </remarks>
    internal class LittleEndianInputStream : ILittleEndianInput
    {
        #region Fields
        private readonly Stream _inputStream;
        #endregion

        #region Constructor
        public LittleEndianInputStream(Stream inputStream)
        {
            _inputStream = inputStream;
        }
        #endregion

        #region Available
        public int Available()
        {
            return (int) (_inputStream.Length - _inputStream.Position);
        }
        #endregion

        #region ReadByte
        /// <summary>
        /// Returns a byte from the stream
        /// </summary>
        /// <returns></returns>
        public int ReadByte()
        {
            return (byte) ReadUByte();
        }
        #endregion

        #region ReadUByte
        /// <summary>
        /// Returns an unsigned byte from the stream
        /// </summary>
        /// <returns></returns>
        public int ReadUByte()
        {
            var b = _inputStream.ReadByte();
            CheckEof(b);
            return b;
        }
        #endregion

        #region ReadDouble
        /// <summary>
        /// Returs a double from the stream
        /// </summary>
        /// <returns></returns>
        public double ReadDouble()
        {
            return BitConverter.Int64BitsToDouble(ReadLong());
        }
        #endregion

        #region ReadInt
        /// <summary>
        /// Returns an integer from the stream
        /// </summary>
        /// <returns></returns>
        public int ReadInt()
        {
            var byte1 = _inputStream.ReadByte();
            var byte2 = _inputStream.ReadByte();
            var byte3 = _inputStream.ReadByte();
            var byte4 = _inputStream.ReadByte();
            CheckEof(byte1 | byte2 | byte3 | byte4);
            return (byte4 << 24) + (byte3 << 16) + (byte2 << 8) + (byte1 << 0);
        }
        #endregion

        #region ReadLong
        /// <summary>
        /// Returns a long from the stream
        /// </summary>
        /// <returns></returns>
        public long ReadLong()
        {
            var byte0 = _inputStream.ReadByte();
            var byte1 = _inputStream.ReadByte();
            var byte2 = _inputStream.ReadByte();
            var byte3 = _inputStream.ReadByte();
            var byte4 = _inputStream.ReadByte();
            var byte5 = _inputStream.ReadByte();
            var byte6 = _inputStream.ReadByte();
            var byte7 = _inputStream.ReadByte();
            CheckEof(byte0 | byte1 | byte2 | byte3 | byte4 | byte5 | byte6 | byte7);
            return (((long) byte7 << 56) +
                    ((long) byte6 << 48) +
                    ((long) byte5 << 40) +
                    ((long) byte4 << 32) +
                    ((long) byte3 << 24) +
                    (byte2 << 16) +
                    (byte1 << 8) +
                    (byte0 << 0));
        }
        #endregion

        #region ReadShort
        /// <summary>
        /// Returns a short from the stream
        /// </summary>
        /// <returns></returns>
        public short ReadShort()
        {
            return (short) ReadUShort();
        }
        #endregion

        #region ReadUShort
        /// <summary>
        /// Returns an unsigned short from the stream
        /// </summary>
        /// <returns></returns>
        public int ReadUShort()
        {
            var byte1 = _inputStream.ReadByte();
            var byte2 = _inputStream.ReadByte();
            CheckEof(byte1 | byte2);
            return (byte2 << 8) + (byte1 << 0);
        }
        #endregion

        #region ReadFully
        public void ReadFully(byte[] buffer)
        {
            ReadFully(buffer, 0, buffer.Length);
        }

        public void ReadFully(byte[] buffer, int offset, int length)
        {
            var max = offset + length;
            for (var i = offset; i < max; i++)
            {
                var b = (byte) _inputStream.ReadByte();
                CheckEof(b);
                buffer[i] = b;
            }
        }
        #endregion

        #region CheckEof
        // ReSharper disable once UnusedParameter.Local
        private static void CheckEof(int value)
        {
            if (value < 0)
                throw new OEFileIsCorrupt("Unexpected end-of-file");
        }
        #endregion
    }
}