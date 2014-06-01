using System;
using System.Collections.Generic;
using System.IO;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{
    /// <summary>
    ///     Stream decorator for a Sector or miniSector chain
    /// </summary>
    internal class StreamView : Stream
    {
        #region Fields
        /// <summary>
        ///     The <see cref="_stream" /> buffer
        /// </summary>
        private readonly byte[] _buf = new byte[4];

        /// <summary>
        ///     The sector chain
        /// </summary>
        private readonly List<Sector> _sectorChain;

        /// <summary>
        ///     The size of a sector in the <see cref="_stream" />
        /// </summary>
        private readonly int _sectorSize;

        /// <summary>
        ///     The stream to view
        /// </summary>
        private readonly Stream _stream;

        /// <summary>
        ///     The length of the <see cref="_stream" />
        /// </summary>
        private long _length;

        /// <summary>
        ///     The current position in the stresm
        /// </summary>
        private long _position;
        #endregion

        #region Properties
        public List<Sector> BaseSectorChain
        {
            get { return _sectorChain; }
        }

        public override bool CanRead
        {
            get { return true; }
        }

        public override bool CanSeek
        {
            get { return true; }
        }

        public override bool CanWrite
        {
            get { return true; }
        }

        public override long Length
        {
            get { return _length; }
        }

        public override long Position
        {
            get { return _position; }

            set
            {
                if (_position > _length - 1)
                    throw new ArgumentOutOfRangeException("value");

                _position = value;
            }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Creates this object
        /// </summary>
        /// <param name="sectorChain"></param>
        /// <param name="sectorSize"></param>
        /// <param name="stream"></param>
        /// <exception cref="CFException">Raised when <see cref="sectorChain"/> is null or <see cref="sectorSize"/> is zero or smaller</exception>
        public StreamView(List<Sector> sectorChain, int sectorSize, Stream stream)
        {
            if (sectorChain == null)
                throw new CFException("Sector chain cannot be null");

            if (sectorSize <= 0)
                throw new CFException("Sector size must be greater than zero");

            _sectorChain = sectorChain;
            _sectorSize = sectorSize;
            _stream = stream;
        }

        public StreamView(List<Sector> sectorChain, int sectorSize, long length, Queue<Sector> availableSectors,
            Stream stream)
            : this(sectorChain, sectorSize, stream)
        {
            AdjustLength(length, availableSectors);
        }

        public StreamView(List<Sector> sectorChain, int sectorSize, long length, Stream stream)
            : this(sectorChain, sectorSize, stream)
        {
            AdjustLength(length);
        }
        #endregion

        #region Flush
        /// <summary>
        /// Flushes the stream... NOT IMPLEMENTED
        /// </summary>
        public override void Flush()
        {
        }
        #endregion

        #region ReadInt32
        /// <summary>
        /// Reads a 32 bit integer from the <see cref="_stream"/>
        /// </summary>
        /// <returns></returns>
        public int ReadInt32()
        {
            Read(_buf, 0, 4);
            return (((_buf[0] | (_buf[1] << 8)) | (_buf[2] << 16)) | (_buf[3] << 24));
        }
        #endregion

        #region Read
        /// <summary>
        /// Reads from the <see cref="_stream"/>
        /// </summary>
        /// <param name="buffer"></param>
        /// <param name="offset"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        public override int Read(byte[] buffer, int offset, int count)
        {
            var read = 0;

            if (_sectorChain == null || _sectorChain.Count <= 0) return 0;
            // First sector
            var secIndex = (int) (_position/_sectorSize);

            // Bytes to read count is the min between request count
            // and sector border

            var toRead = Math.Min(
                _sectorChain[0].Size - ((int) _position%_sectorSize),
                count);

            if (secIndex < _sectorChain.Count)
            {
                Buffer.BlockCopy(
                    _sectorChain[secIndex].GetData(),
                    (int) (_position%_sectorSize),
                    buffer,
                    offset,
                    toRead
                    );
            }

            read += toRead;

            secIndex++;

            // Central sectors
            while (read < (count - _sectorSize))
            {
                toRead = _sectorSize;

                Buffer.BlockCopy(
                    _sectorChain[secIndex].GetData(),
                    0,
                    buffer,
                    offset + read,
                    toRead
                    );

                read += toRead;
                secIndex++;
            }

            // Last sector
            toRead = count - read;

            if (toRead != 0)
            {
                Buffer.BlockCopy(
                    _sectorChain[secIndex].GetData(),
                    0,
                    buffer,
                    offset + read,
                    toRead
                    );

                read += toRead;
            }

            _position += read;
            return read;
        }
        #endregion

        #region Seek
        /// <summary>
        /// Seeks a new position in the <see cref="_stream"/> from the <see cref="origin"/>
        /// </summary>
        /// <param name="offset"></param>
        /// <param name="origin"></param>
        /// <returns></returns>
        public override long Seek(long offset, SeekOrigin origin)
        {
            switch (origin)
            {
                case SeekOrigin.Begin:
                    _position = offset;
                    break;

                case SeekOrigin.Current:
                    _position += offset;
                    break;

                case SeekOrigin.End:
                    _position = Length - offset;
                    break;
            }

            AdjustLength(_position);

            return _position;
        }
        #endregion

        #region AdjustLength
        /// <summary>
        /// Adjusts the length of the <see cref="_stream"/>
        /// </summary>
        /// <param name="value"></param>
        private void AdjustLength(long value)
        {
            AdjustLength(value, null);
        }

        /// <summary>
        /// Adjusts the length of the <see cref="_stream"/>
        /// </summary>
        /// <param name="value"></param>
        /// <param name="availableSectors"></param>
        private void AdjustLength(long value, Queue<Sector> availableSectors)
        {
            _length = value;

            var delta = value - (_sectorChain.Count*(long) _sectorSize);

            if (delta <= 0) return;

            var nextSector = (int) Math.Ceiling(((double) delta/_sectorSize));

            while (nextSector > 0)
            {
                Sector sector;

                if (availableSectors == null || availableSectors.Count == 0)
                    sector = new Sector(_sectorSize, _stream);
                else
                    sector = availableSectors.Dequeue();

                _sectorChain.Add(sector);
                nextSector--;
            }
        }

        /// <summary>
        /// Adjusts the length of the <see cref="_stream"/>
        /// </summary>
        /// <param name="value"></param>
        public override void SetLength(long value)
        {
            AdjustLength(value);
        }
        #endregion

        #region Write
        /// <summary>
        /// Writes the <see cref="buffer"/> to the <see cref="_stream"/> on the <see cref="offset"/>
        /// </summary>
        /// <param name="buffer"></param>
        /// <param name="offset"></param>
        /// <param name="count"></param>
        public override void Write(byte[] buffer, int offset, int count)
        {
            var byteWritten = 0;

            // Assure length
            if ((_position + count) > _length)
                AdjustLength((_position + count));

            if (_sectorChain == null) return;
            // First sector
            var secOffset = (int) (_position/_sectorSize);
            var secShift = (int) _position%_sectorSize;

            var roundByteWritten = Math.Min(_sectorSize - (int) (_position%_sectorSize), count);

            if (secOffset < _sectorChain.Count)
            {
                Buffer.BlockCopy(
                    buffer,
                    offset,
                    _sectorChain[secOffset].GetData(),
                    secShift,
                    roundByteWritten
                    );

                _sectorChain[secOffset].DirtyFlag = true;
            }

            byteWritten += roundByteWritten;
            offset += roundByteWritten;
            secOffset++;

            // Central sectors
            while (byteWritten < (count - _sectorSize))
            {
                roundByteWritten = _sectorSize;

                Buffer.BlockCopy(
                    buffer,
                    offset,
                    _sectorChain[secOffset].GetData(),
                    0,
                    roundByteWritten
                    );

                _sectorChain[secOffset].DirtyFlag = true;

                byteWritten += roundByteWritten;
                offset += roundByteWritten;
                secOffset++;
            }

            // Last sector
            roundByteWritten = count - byteWritten;

            if (roundByteWritten != 0)
            {
                Buffer.BlockCopy(
                    buffer,
                    offset,
                    _sectorChain[secOffset].GetData(),
                    0,
                    roundByteWritten
                    );

                _sectorChain[secOffset].DirtyFlag = true;
            }

            _position += count;
        }
        #endregion
    }
}