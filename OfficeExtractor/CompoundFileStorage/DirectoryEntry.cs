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
        #endregion

        #region Properties
        public byte[] EntryName { get; private set; }
        public ushort NameLength { get; private set; }

        public StgType StgType { get; set; }

        public StgColor StgColor { get; set; }

        public int LeftSibling { get; set; }

        public int RightSibling { get; set; }

        public int Child { get; set; }

        public Guid StorageCLSID { get; set; }


        public int StateBits { get; set; }

        public byte[] CreationDate { get; set; }

        public byte[] ModifyDate { get; set; }

        public int StartSector { get; set; }

        public long Size { get; set; }

        public int SID { get; set; }
        #endregion

        #region Constructor
        public DirectoryEntry(StgType stgType)
        {
            StgColor = StgColor.Black;
            SID = -1;
            EntryName = new byte[64];
            RightSibling = Nostream;
            LeftSibling = Nostream;
            Child = Nostream;
            StorageCLSID = Guid.NewGuid();
            CreationDate = new byte[] {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
            ModifyDate = new byte[] {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
            StartSector = Sector.Endofchain;
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
        ///     Gets the name of the directory entry
        /// </summary>
        /// <returns></returns>
        public string GetEntryName()
        {
            if (EntryName != null && EntryName.Length > 0)
            {
                return Encoding.Unicode.GetString(EntryName).Remove((NameLength - 1)/2);
            }
            return string.Empty;
        }
        #endregion

        #region SetEntryName
        /// <summary>
        ///     Sets the name of the directory entry
        /// </summary>
        /// <param name="entryName"></param>
        /// <exception cref="CFException">Raised when an invalid character is used or the length is longer then 31</exception>
        public void SetEntryName(string entryName)
        {
            if (
                entryName.Contains(@"\") ||
                entryName.Contains(@"/") ||
                entryName.Contains(@":") ||
                entryName.Contains(@"!")
                )
                throw new CFException(
                    "Invalid character in entry: the characters '\\', '/', ':','!' cannot be used in entry name");

            if (entryName.Length > 31)
                throw new CFException("Entry name MUST be smaller than 31 characters");


            var temp = Encoding.Unicode.GetBytes(entryName);
            var newName = new byte[64];
            Buffer.BlockCopy(temp, 0, newName, 0, temp.Length);
            newName[temp.Length] = 0x00;
            newName[temp.Length + 1] = 0x00;

            EntryName = newName;
            NameLength = (ushort) (temp.Length + 2);
        }
        #endregion

        #region CompareTo
        /// <summary>
        ///     Compares the object to the current IDirectory entry
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            const int thisIsGreater = 1;
            const int otherIsGreater = -1;
            var otherDir = obj as IDirectoryEntry;

            if (otherDir == null)
                throw new CFException("Invalid casting: compared object does not implement IDirectorEntry interface");

            if (NameLength > otherDir.NameLength)
            {
                return thisIsGreater;
            }
            if (NameLength < otherDir.NameLength)
            {
                return otherIsGreater;
            }
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

        #region Equals
        public override bool Equals(object obj)
        {
            return CompareTo(obj) == 0;
        }
        #endregion

        #region GetHashCode
        /// <summary>
        ///     FNV hash, short for Fowler/Noll/Vo
        /// </summary>
        /// <param name="buffer"></param>
        /// <returns>(not warranted) unique hash for byte array</returns>
        private static ulong fnv_hash(IList<byte> buffer)
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
            return (int) fnv_hash(EntryName);
        }
        #endregion

        #region Write
        /// <summary>
        ///     Writes to the <see cref="stream" />
        /// </summary>
        /// <param name="stream"></param>
        public void Write(Stream stream)
        {
            var streamRw = new StreamRW(stream);

            streamRw.Write(EntryName);
            streamRw.Write(NameLength);
            streamRw.Write((byte) StgType);
            streamRw.Write((byte) StgColor);
            streamRw.Write(LeftSibling);
            streamRw.Write(RightSibling);
            streamRw.Write(Child);
            streamRw.Write(StorageCLSID.ToByteArray());
            streamRw.Write(StateBits);
            streamRw.Write(CreationDate);
            streamRw.Write(ModifyDate);
            streamRw.Write(StartSector);
            streamRw.Write(Size);

            streamRw.Close();
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

            EntryName = streamRw.ReadBytes(64);
            NameLength = streamRw.ReadUInt16();
            StgType = (StgType) streamRw.ReadByte();
            streamRw.ReadByte();
            LeftSibling = streamRw.ReadInt32();
            RightSibling = streamRw.ReadInt32();
            Child = streamRw.ReadInt32();

            if (StgType == StgType.StgInvalid)
            {
                LeftSibling = Nostream;
                RightSibling = Nostream;
                Child = Nostream;
            }

            StorageCLSID = new Guid(streamRw.ReadBytes(16));
            StateBits = streamRw.ReadInt32();
            CreationDate = streamRw.ReadBytes(8);
            ModifyDate = streamRw.ReadBytes(8);
            StartSector = streamRw.ReadInt32();
            Size = streamRw.ReadInt64();
        }
        #endregion

        #region Name
        /// <summary>
        ///     The name of the directory entry
        /// </summary>
        public string Name
        {
            get { return GetEntryName(); }
        }
        #endregion
    }
}