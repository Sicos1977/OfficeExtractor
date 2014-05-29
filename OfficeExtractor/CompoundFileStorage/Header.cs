using System.IO;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{
    /// <summary>
    ///     This class contains the header of an OLE structured storage file
    /// </summary>
    internal class Header
    {
        #region Fields
        //0 8 Compound document file identifier: D0H CFH 11H E0H A1H B1H 1AH E1H
        /// <summary>
        ///     Structured Storage signature
        /// </summary>
        private readonly byte[] _oleCFSSignature = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};

        private byte[] _headerSignature = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};

        //8 16 Unique identifier (UID) of this file (not of interest in the following, may be all 0)
        private ushort _majorVersion = 0x0003;

        //24 2 Revision number of the file format (most used is 003EH)
        private ushort _minorVersion = 0x003E;

        public Header() : this(3)
        {
        }
        #endregion

        #region Properties
        /// <summary>
        ///     The signature of the header
        /// </summary>
        public byte[] HeaderSignature
        {
            get { return _headerSignature; }
        }

        /// <summary>
        ///     The CLSID of the header
        /// </summary>
        public byte[] CLSID { get; set; }

        /// <summary>
        ///     Minor version number of the header
        /// </summary>
        public ushort MinorVersion
        {
            get { return _minorVersion; }
        }

        /// <summary>
        ///     Version number of the file format (most used is 0003H)
        /// </summary>
        public ushort MajorVersion
        {
            get { return _majorVersion; }
        }

        /// <summary>
        ///     Byte order identifier (➜4.2): FEH FFH = Little-Endian FFH FEH = Big-Endian
        /// </summary>
        public ushort ByteOrder { get; private set; }

        /// <summary>
        ///     Size of a sector in the compound document file (➜3.1) in power-of-two (ssz), real sector
        ///     size is sec_size = 2ssz bytes (minimum value is 7 which means 128 bytes, most used
        ///     value is 9 which means 512 bytes)
        /// </summary>
        public ushort SectorShift { get; private set; }

        /// <summary>
        ///     Size of a short-sector in the short-stream container stream (➜6.1) in power-of-two (sssz),
        ///     real short-sector size is short_sec_size = 2sssz bytes (maximum value is sector size
        ///     ssz, see above, most used value is 6 which means 64 bytes)
        /// </summary>
        public ushort MiniSectorShift { get; private set; }

        /// <summary>
        ///     Not used
        /// </summary>
        public byte[] UnUsed { get; private set; }

        /// <summary>
        ///     Total number of sectors used Directory (➜5.2)
        /// </summary>
        public int DirectorySectorsNumber { get; set; }

        /// <summary>
        ///     Total number of sectors used for the sector allocation table (➜5.2)
        /// </summary>
        public int FATSectorsNumber { get; set; }

        /// <summary>
        ///     SecID of first sector of the directory stream (➜7)
        /// </summary>
        public int FirstDirectorySectorId { get; set; }

        /// <summary>
        ///     Not used
        /// </summary>
        public uint UnUsed2 { get; private set; }

        /// <summary>
        ///     Minimum size of a standard stream (in bytes, minimum allowed and most used size is 4096
        ///     bytes), streams with an actual size smaller than (and not equal to) this value are stored as
        ///     short-streams (➜6)
        /// </summary>
        public uint MinSizeStandardStream { get; set; }

        /// <summary>
        ///     This integer field contains the starting sector number for the mini FAT
        /// </summary>
        public int FirstMiniFATSectorId { get; set; }

        /// <summary>
        ///     Total number of sectors used for the short-sector allocation table (➜6.2)
        /// </summary>
        public uint MiniFATSectorsNumber { get; set; }

        /// <summary>
        ///     SecID of first sector of the master sector allocation table (➜5.1), or –2
        ///     (End Of Chain //SecID, ➜3.1) if no additional sectors used
        /// </summary>
        public int FirstDIFATSectorId { get; set; }

        /// <summary>
        ///     Total number of sectors used for the master sector allocation table (➜5.1)
        /// </summary>
        public uint DIFATSectorsNumber { get; set; }

        /// <summary>
        ///     First part of the master sector allocation table (➜5.1) containing 109 SecIDs
        /// </summary>
        public int[] DIFAT { get; private set; }
        #endregion

        #region Constructor
        /// <summary>
        /// Creates this object
        /// </summary>
        /// <param name="version"></param>
        /// <exception cref="CFException">Raised when an invalid <see cref="version"/> is used</exception>
        public Header(ushort version)
        {
            DIFAT = new int[109];
            UnUsed = new byte[6];
            CLSID = new byte[16];
            FirstDIFATSectorId = Sector.Endofchain;
            FirstMiniFATSectorId = unchecked((int) 0xFFFFFFFE);
            MinSizeStandardStream = 4096;
            FirstDirectorySectorId = Sector.Endofchain;
            ByteOrder = 0xFFFE;
            SectorShift = 9;
            MiniSectorShift = 6;
            switch (version)
            {
                case 3:
                    _majorVersion = 3;
                    SectorShift = 0x0009;
                    break;

                case 4:
                    _majorVersion = 4;
                    SectorShift = 0x000C;
                    break;

                default:
                    throw new CFException("Invalid Compound File Format version");
            }

            for (var i = 0; i < 109; i++)
            {
                DIFAT[i] = Sector.FreeSector;
            }
        }
        #endregion

        #region Read
        /// <summary>
        ///     Reads the header from the stream
        /// </summary>
        /// <param name="stream"></param>
        public void Read(Stream stream)
        {
            var rw = new StreamReader(stream);

            _headerSignature = rw.ReadBytes(8);
            CheckSignature();
            CLSID = rw.ReadBytes(16);
            _minorVersion = rw.ReadUInt16();
            _majorVersion = rw.ReadUInt16();
            CheckVersion();
            ByteOrder = rw.ReadUInt16();
            SectorShift = rw.ReadUInt16();
            MiniSectorShift = rw.ReadUInt16();
            UnUsed = rw.ReadBytes(6);
            DirectorySectorsNumber = rw.ReadInt32();
            FATSectorsNumber = rw.ReadInt32();
            FirstDirectorySectorId = rw.ReadInt32();
            UnUsed2 = rw.ReadUInt32();
            MinSizeStandardStream = rw.ReadUInt32();
            FirstMiniFATSectorId = rw.ReadInt32();
            MiniFATSectorsNumber = rw.ReadUInt32();
            FirstDIFATSectorId = rw.ReadInt32();
            DIFATSectorsNumber = rw.ReadUInt32();

            for (var i = 0; i < 109; i++)
                DIFAT[i] = rw.ReadInt32();

            rw.Close();
        }
        #endregion

        #region CheckVersion
        /// <summary>
        ///     Checks if the file has a valid OLE structured storage version number, only 3 and 4 are supported
        /// </summary>
        /// <exception cref="CFFormatException">Raised when the compound file storage contains an invalid format</exception>
        private void CheckVersion()
        {
            if (_majorVersion != 3 && _majorVersion != 4)
                throw new CFFormatException(
                    "Unsupported binary file format version, only support for compound files with major version equal to 3 or 4 ");
        }
        #endregion

        #region CheckSignature
        /// <summary>
        ///     Checks if the file has a valid OLE structured storage signature
        /// </summary>
        /// <exception cref="CFFormatException">Raised when the file is invalid</exception>
        private void CheckSignature()
        {
            for (var i = 0; i < _headerSignature.Length; i++)
            {
                if (_headerSignature[i] != _oleCFSSignature[i])
                    throw new CFFormatException("Invalid OLE structured storage file");
            }
        }
        #endregion
    }
}