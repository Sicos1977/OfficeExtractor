using System;
using System.IO;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Interfaces
{
    /// <summary>
    ///     The directory entry interface
    /// </summary>
    public interface IDirectoryEntry : IComparable
    {
        int Child { get; set; }

        byte[] CreationDate { get; set; }

        byte[] EntryName { get; }

        int LeftSibling { get; set; }

        byte[] ModifyDate { get; set; }

        string Name { get; }

        ushort NameLength { get; }

        int RightSibling { get; set; }

        int SID { get; set; }

        long Size { get; set; }

        int StartSector { get; set; }

        int StateBits { get; set; }

        StgColor StgColor { get; set; }

        StgType StgType { get; set; }

        Guid StorageCLSID { get; set; }
        string GetEntryName();
        void Read(Stream stream);
        void SetEntryName(string entryName);

        void Write(Stream stream);
    }
}