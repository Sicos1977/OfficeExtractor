using System.IO;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{
    /// <summary>
    ///     Used to read from the stream
    /// </summary>
    internal class StreamRW
    {
        #region Fields
        private readonly byte[] _buffer = new byte[8];
        private readonly Stream _stream;
        #endregion

        #region Constructor
        public StreamRW(Stream stream)
        {
            _stream = stream;
        }
        #endregion

        #region Seek
        public long Seek(long offset)
        {
            return _stream.Seek(offset, SeekOrigin.Begin);
        }
        #endregion

        #region ReadByte
        public byte ReadByte()
        {
            return (byte) _stream.ReadByte();
        }
        #endregion

        #region ReadUInt16
        public ushort ReadUInt16()
        {
            _stream.Read(_buffer, 0, 2);
            return (ushort) (_buffer[0] | (_buffer[1] << 8));
        }
        #endregion

        #region ReadInt32
        public int ReadInt32()
        {
            _stream.Read(_buffer, 0, 4);
            return (_buffer[0] | (_buffer[1] << 8) | (_buffer[2] << 16) | (_buffer[3] << 24));
        }
        #endregion

        #region ReadUInt32
        public uint ReadUInt32()
        {
            _stream.Read(_buffer, 0, 4);
            return (uint) (_buffer[0] | (_buffer[1] << 8) | (_buffer[2] << 16) | (_buffer[3] << 24));
        }
        #endregion

        #region ReadInt64
        public long ReadInt64()
        {
            _stream.Read(_buffer, 0, 8);
            var ls = (uint) (_buffer[0] | (_buffer[1] << 8) | (_buffer[2] << 16) | (_buffer[3] << 24));
            var ms = (uint) ((_buffer[4]) | (_buffer[5] << 8) | (_buffer[6] << 16) | (_buffer[7] << 24));
            return ((ms << 32) | ls);
        }
        #endregion

        #region ReadUInt64
        public ulong ReadUInt64()
        {
            _stream.Read(_buffer, 0, 8);
            return
                (ulong)
                    (_buffer[0] | (_buffer[1] << 8) | (_buffer[2] << 16) | (_buffer[3] << 24) | (_buffer[4] << 32) |
                     (_buffer[5] << 40) | (_buffer[6] << 48) | (_buffer[7] << 56));
        }
        #endregion

        #region ReadBytes
        public byte[] ReadBytes(int count)
        {
            var result = new byte[count];
            _stream.Read(result, 0, count);
            return result;
        }

        public byte[] ReadBytes(int count, out int read)
        {
            var result = new byte[count];
            read = _stream.Read(result, 0, count);
            return result;
        }
        #endregion

        #region Write
        public void Write(byte b)
        {
            _stream.WriteByte(b);
        }

        public void Write(ushort value)
        {
            _buffer[0] = (byte) value;
            _buffer[1] = (byte) (value >> 8);

            _stream.Write(_buffer, 0, 2);
        }

        public void Write(int value)
        {
            _buffer[0] = (byte) value;
            _buffer[1] = (byte) (value >> 8);
            _buffer[2] = (byte) (value >> 16);
            _buffer[3] = (byte) (value >> 24);

            _stream.Write(_buffer, 0, 4);
        }

        public void Write(long value)
        {
            _buffer[0] = (byte) value;
            _buffer[1] = (byte) (value >> 8);
            _buffer[2] = (byte) (value >> 16);
            _buffer[3] = (byte) (value >> 24);
            _buffer[4] = (byte) (value >> 32);
            _buffer[5] = (byte) (value >> 40);
            _buffer[6] = (byte) (value >> 48);
            _buffer[7] = (byte) (value >> 56);

            _stream.Write(_buffer, 0, 8);
        }

        public void Write(uint value)
        {
            _buffer[0] = (byte) value;
            _buffer[1] = (byte) (value >> 8);
            _buffer[2] = (byte) (value >> 16);
            _buffer[3] = (byte) (value >> 24);

            _stream.Write(_buffer, 0, 4);
        }

        public void Write(byte[] value)
        {
            _stream.Write(value, 0, value.Length);
        }
        #endregion

        #region Close
        public void Close()
        {
            //Nothing to do ;-)
        }
        #endregion
    }
}