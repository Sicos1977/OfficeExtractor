using System;
using System.IO;
using OfficeExtractor.Biff8.Interfaces;

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

namespace OfficeExtractor.Biff8
{
    /// <summary>
    ///     Wraps an <see cref="Stream" /> providing <see cref="ILittleEndianOutput" />
    /// </summary>
    internal class LittleEndianOutputStream : ILittleEndianOutput, IDisposable
    {
        #region Fields
        private Stream _outputStream;
        #endregion

        #region Constructor
        public LittleEndianOutputStream(Stream output)
        {
            _outputStream = output;
        }
        #endregion

        #region WriteByte
        /// <summary>
        /// Writes a byte to the <see cref="_outputStream"/>
        /// </summary>
        /// <param name="b"></param>
        public void WriteByte(int b)
        {
            _outputStream.WriteByte((byte) b);
        }
        #endregion

        #region WriteDouble
        /// <summary>
        /// Writes a double to the <see cref="_outputStream"/>
        /// </summary>
        /// <param name="d"></param>
        public void WriteDouble(double d)
        {
            WriteLong(BitConverter.DoubleToInt64Bits(d));
        }
        #endregion

        #region WriteInt
        /// <summary>
        /// Writes an integer to the <see cref="_outputStream"/>
        /// </summary>
        /// <param name="i"></param>
        public void WriteInt(int i)
        {
            var byte3 = (i >> 24) & 0xFF;
            var byte2 = (i >> 16) & 0xFF;
            var byte1 = (i >> 8) & 0xFF;
            var byte0 = (i >> 0) & 0xFF;
            _outputStream.WriteByte((byte) byte0);
            _outputStream.WriteByte((byte) byte1);
            _outputStream.WriteByte((byte) byte2);
            _outputStream.WriteByte((byte) byte3);
        }
        #endregion

        #region WriteLong
        /// <summary>
        /// Writes a long to the <see cref="_outputStream"/>
        /// </summary>
        /// <param name="l"></param>
        public void WriteLong(long l)
        {
            WriteInt((int) (l >> 0));
            WriteInt((int) (l >> 32));
        }
        #endregion

        #region WriteShort
        /// <summary>
        /// Writes a short to the <see cref="_outputStream"/>
        /// </summary>
        /// <param name="s"></param>
        public void WriteShort(int s)
        {
            var byte1 = (s >> 8) & 0xFF;
            var byte0 = (s >> 0) & 0xFF;
            _outputStream.WriteByte((byte) byte0);
            _outputStream.WriteByte((byte) byte1);
        }
        #endregion

        #region Write
        /// <summary>
        /// Writes a byte array to the <see cref="_outputStream"/>
        /// </summary>
        /// <param name="b"></param>
        public void Write(byte[] b)
        {
            _outputStream.Write(b, 0, b.Length);
        }

        /// <summary>
        /// Writes tot the <see cref="_outputStream"/>
        /// </summary>
        /// <param name="b"></param>
        /// <param name="offset"></param>
        /// <param name="length"></param>
        public void Write(byte[] b, int offset, int length)
        {
            _outputStream.Write(b, offset, length);
        }
        #endregion

        #region Flush
        /// <summary>
        /// Flushes the <see cref="_outputStream"/>
        /// </summary>
        public void Flush()
        {
            _outputStream.Flush();
        }
        #endregion

        #region Dispose methods
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (!disposing) return;
            if (null == _outputStream) return;
            _outputStream.Dispose();
            _outputStream = null;
        }
        #endregion
    }
}