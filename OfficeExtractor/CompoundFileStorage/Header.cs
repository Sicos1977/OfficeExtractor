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
        /// <summary>
        ///     Structured Storage signature
        /// </summary>
        private readonly byte[] _oleCFSSignature = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};
        #endregion

        #region Properties
        /// <summary>
        ///     Compound document file identifier: D0H CFH 11H E0H A1H B1H 1AH E1H
        /// </summary>
        public byte[] HeaderSignature { get; private set; }

        /// <summary>
        ///     Unique identifier (UID) of this file (not of interest in the following, may be all 0)
        /// </summary>
        public byte[] CLSID { get; set; }

        /// <summary>
        ///     Revision number of the file format (most used is 003EH)
        /// </summary>
        public ushort MinorVersion { get; private set; }

        /// <summary>
        ///     Version number of the file format (most used is 0003H)
        /// </summary>
        public ushort MajorVersion { get; private set; }

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

        //60 4 SecID of first sector of the short-sector allocation table (➜6.2), or –2 (End Of Chain
        //SecID, ➜3.1) if not extant

        /// <summary>
        ///     This integer field contains the starting sector number for the mini FAT
        /// </summary>
        public int FirstMiniFATSectorId { get; set; }

        /// <summary>
        ///     Total number of sectors used for the short-sector allocation table (➜6.2)
        /// </summary>
        public uint MiniFATSectorsNumber { get; set; }

        /// <summary>
        ///     SecID of first sector of the master sector allocation table (➜5.1), or –2 (End Of Chain
        ///     SecID, ➜3.1) if no additional sectors used
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

        #region Constructors
        public Header() : this(3)
        {
        }

        public Header(ushort version)
        {
            HeaderSignature = _oleCFSSignature;
            DIFAT = new int[109];
            UnUsed = new byte[6];
            MiniSectorShift = 6;
            SectorShift = 9;
            ByteOrder = 0xFFFE;
            MajorVersion = 0x0003;
            MinorVersion = 0x003E;
            CLSID = new byte[16];
            FirstDirectorySectorId = Sector.Endofchain;
            MinSizeStandardStream = 4096;
            FirstMiniFATSectorId = unchecked((int) 0xFFFFFFFE);
            FirstDIFATSectorId = Sector.Endofchain;
            
            switch (version)
            {
                case 3:
                    MajorVersion = 3;
                    SectorShift = 0x0009;
                    break;

                case 4:
                    MajorVersion = 4;
                    SectorShift = 0x000C;
                    break;

                default:
                    throw new CFException("Invalid Compound File Format version");
            }

            for (var i = 0; i < 109; i++)
                DIFAT[i] = Sector.FreeSector;
        }
        #endregion

        #region Read
        /// <summary>
        ///     Reads from the <see cref="stream" />
        /// </summary>
        /// <param name="stream"></param>
        public void Read(Stream stream)
        {
            var streamRw = new StreamRW(stream);

            HeaderSignature = streamRw.ReadBytes(8);
            CheckSignature();
            CLSID = streamRw.ReadBytes(16);
            MinorVersion = streamRw.ReadUInt16();
            MajorVersion = streamRw.ReadUInt16();
            CheckVersion();
            ByteOrder = streamRw.ReadUInt16();
            SectorShift = streamRw.ReadUInt16();
            MiniSectorShift = streamRw.ReadUInt16();
            UnUsed = streamRw.ReadBytes(6);
            DirectorySectorsNumber = streamRw.ReadInt32();
            FATSectorsNumber = streamRw.ReadInt32();
            FirstDirectorySectorId = streamRw.ReadInt32();
            UnUsed2 = streamRw.ReadUInt32();
            MinSizeStandardStream = streamRw.ReadUInt32();
            FirstMiniFATSectorId = streamRw.ReadInt32();
            MiniFATSectorsNumber = streamRw.ReadUInt32();
            FirstDIFATSectorId = streamRw.ReadInt32();
            DIFATSectorsNumber = streamRw.ReadUInt32();

            for (var i = 0; i < 109; i++)
                DIFAT[i] = streamRw.ReadInt32();

            streamRw.Close();
        }
        #endregion

        #region Write
        /// <summary>
        ///     Writes to the <see cref="stream"/>
        /// </summary>
        /// <param name="stream"></param>
        public void Write(Stream stream)
        {
            var streamRw = new StreamRW(stream);

            streamRw.Write(HeaderSignature);
            streamRw.Write(CLSID);
            streamRw.Write(MinorVersion);
            streamRw.Write(MajorVersion);
            streamRw.Write(ByteOrder);
            streamRw.Write(SectorShift);
            streamRw.Write(MiniSectorShift);
            streamRw.Write(UnUsed);
            streamRw.Write(DirectorySectorsNumber);
            streamRw.Write(FATSectorsNumber);
            streamRw.Write(FirstDirectorySectorId);
            streamRw.Write(UnUsed2);
            streamRw.Write(MinSizeStandardStream);
            streamRw.Write(FirstMiniFATSectorId);
            streamRw.Write(MiniFATSectorsNumber);
            streamRw.Write(FirstDIFATSectorId);
            streamRw.Write(DIFATSectorsNumber);

            foreach (var i in DIFAT)
                streamRw.Write(i);

            if (MajorVersion == 4)
            {
                var zeroHead = new byte[3584];
                streamRw.Write(zeroHead);
            }

            streamRw.Close();
        }
        #endregion
        
        #region CheckVersion
        private void CheckVersion()
        {
            if (MajorVersion != 3 && MajorVersion != 4)
                throw new CFFileFormatException(
                    "Unsupported binary file format version there is only support for compound files with major version equal to 3 or 4");
        }
        #endregion

        #region CheckSignature
        private void CheckSignature()
        {
            for (var i = 0; i < HeaderSignature.Length; i++)
            {
                if (HeaderSignature[i] != _oleCFSSignature[i])
                    throw new CFFileFormatException("Invalid OLE structured storage file");
            }
        }
        #endregion
    }
}