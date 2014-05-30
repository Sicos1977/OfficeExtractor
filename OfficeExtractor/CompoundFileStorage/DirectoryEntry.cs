using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Interfaces;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{

    #region Enum StgType
    public enum StgType
    {
        StgInvalid = 0,
        StgStorage = 1,
        StgStream = 2,
        StgLockbytes = 3,
        StgProperty = 4,
        StgRoot = 5
    }
    #endregion

    #region Enum StgColor
    public enum StgColor
    {
        Red = 0,
        Black = 1
    }
    #endregion

    internal class DirectoryEntry : IDirectoryEntry
    {
        #region Fields
        internal static Int32 Nostream = unchecked((int) 0xFFFFFFFF);
        private byte[] _entryName = new byte[64];
        private ushort _nameLength;
        #endregion

        #region Properties
        public int SID { get; set; }

        public byte[] EntryName
        {
            get { return _entryName; }
        }

        public ushort NameLength
        {
            get { return _nameLength; }
            set { throw new NotImplementedException(); }
        }

        public StgType StgType { get; set; }

        public int LeftSibling { get; set; }

        public int RightSibling { get; set; }

        public int Child { get; set; }

        public Guid StorageCLSID { get; set; }

        public int StateBits { get; set; }

        public byte[] CreationDate { get; set; }

        public byte[] ModifyDate { get; set; }

        public int StartSector { get; set; }

        public long Size { get; set; }

        public string Name
        {
            get { return GetEntryName(); }
        }
        #endregion

        #region Constructor
        public DirectoryEntry(StgType stgType)
        {
            SID = -1;
            StartSector = Sector.Endofchain;
            ModifyDate = new byte[] {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
            CreationDate = new byte[] {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
            StorageCLSID = Guid.NewGuid();
            Child = Nostream;
            LeftSibling = Nostream;
            RightSibling = Nostream;
            StgType = stgType;

            switch (stgType)
            {
                case StgType.StgStream:

                    StorageCLSID = new Guid("00000000000000000000000000000000");
                    CreationDate = new byte[] {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
                    ModifyDate = new byte[] {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
                    break;

                case StgType.StgStorage:
                    CreationDate = BitConverter.GetBytes((DateTime.Now.ToFileTime()));
                    break;

                case StgType.StgRoot:
                    CreationDate = new byte[] {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
                    ModifyDate = new byte[] {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
                    break;
            }
        }
        #endregion

        #region GetEntryName
        /// <summary>
        /// Returns the entry name
        /// </summary>
        /// <returns></returns>
        public string GetEntryName()
        {
            if (_entryName != null && _entryName.Length > 0)
                return Encoding.Unicode.GetString(_entryName).Remove((_nameLength - 1)/2).Trim();

            return string.Empty;
        }
        #endregion

        #region SetEntryName
        /// <summary>
        /// Sets the entry name
        /// </summary>
        /// <param name="entryName"></param>
        /// <exception cref="CFException">Raised when the <see cref="entryName"/> contains invalid characters or is longer then 31</exception>
        public void SetEntryName(string entryName)
        {
            if (entryName.Contains(@"\") ||
                entryName.Contains(@"/") ||
                entryName.Contains(@":") ||
                entryName.Contains(@"!"))
                throw new CFException(
                    "Invalid character in entry: the characters '\\', '/', ':','!' cannot be used in entry name");

            if (entryName.Length > 31)
                throw new CFException("Entry name MUST be smaller than 31 characters");

            var temp = Encoding.Unicode.GetBytes(entryName);
            var newName = new byte[64];
            Buffer.BlockCopy(temp, 0, newName, 0, temp.Length);
            newName[temp.Length] = 0x00;
            newName[temp.Length + 1] = 0x00;

            _entryName = newName;
            _nameLength = (ushort) (temp.Length + 2);
        }
        #endregion

        #region Equals
        public override bool Equals(object obj)
        {
            return CompareTo(obj) == 0;
        }
        #endregion

        #region CompareTo
        /// <summary>
        /// Compares an object to the current directory entry
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <exception cref="CFException">Raised when the <see cref="obj"/> does not contain an <see cref="IDirectoryEntry"/> interface</exception>
        public int CompareTo(object obj)
        {
            const int thisIsGreater = 1;
            const int otherIsGreater = -1;
            var otherDir = obj as IDirectoryEntry;

            if (otherDir == null)
                throw new CFException("Invalid casting: compared object does not implement IDirectorEntry interface");

            if (NameLength > otherDir.NameLength)
                return thisIsGreater;

            if (NameLength < otherDir.NameLength)
                return otherIsGreater;

            var thisName = Encoding.Unicode.GetString(EntryName, 0, NameLength).ToUpper(CultureInfo.InvariantCulture);
            var otherName =
                Encoding.Unicode.GetString(otherDir.EntryName, 0, otherDir.NameLength)
                    .ToUpper(CultureInfo.InvariantCulture);

            for (var z = 0; z < thisName.Length; z++)
            {
                if (BitConverter.ToInt16(BitConverter.GetBytes(thisName[z]), 0) >
                    BitConverter.ToInt16(BitConverter.GetBytes(otherName[z]), 0))
                    return thisIsGreater;
                if (BitConverter.ToInt16(BitConverter.GetBytes(thisName[z]), 0) <
                    BitConverter.ToInt16(BitConverter.GetBytes(otherName[z]), 0))
                    return otherIsGreater;
            }

            return 0;
        }
        #endregion

        #region Read
        /// <summary>
        /// Read the <see cref="stream"/>
        /// </summary>
        /// <param name="stream"></param>
        public void Read(Stream stream)
        {
            var rw = new StreamReader(stream);

            _entryName = rw.ReadBytes(64);
            _nameLength = rw.ReadUInt16();
            StgType = (StgType) rw.ReadByte();
            rw.ReadByte(); //Ignore color, only black tree
            LeftSibling = rw.ReadInt32();
            RightSibling = rw.ReadInt32();
            Child = rw.ReadInt32();

            if (StgType == StgType.StgInvalid)
            {
                LeftSibling = Nostream;
                RightSibling = Nostream;
                Child = Nostream;
            }

            StorageCLSID = new Guid(rw.ReadBytes(16));
            StateBits = rw.ReadInt32();
            CreationDate = rw.ReadBytes(8);
            ModifyDate = rw.ReadBytes(8);
            StartSector = rw.ReadInt32();
            Size = rw.ReadInt64();
        }
        #endregion

        #region GetHashCode
        /// <summary>
        ///     FNV hash, short for Fowler/Noll/Vo
        /// </summary>
        /// <param name="buffer"></param>
        /// <returns>(not warranted) unique hash for byte array</returns>
        private static ulong FnvHash(IList<byte> buffer)
        {
            ulong h = 2166136261;
            int i;

            for (i = 0; i < buffer.Count; i++)
                h = (h*16777619) ^ buffer[i];

            return h;
        }

        public override int GetHashCode()
        {
            // ReSharper disable once NonReadonlyFieldInGetHashCode
            return (int) FnvHash(_entryName);
        }
        #endregion
    }
}