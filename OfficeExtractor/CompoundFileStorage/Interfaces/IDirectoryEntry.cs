using System;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Interfaces
{
    /// <summary>
    /// The directory entry interface
    /// </summary>
    public interface IDirectoryEntry : IComparable
    {
        int Child { get; set; }

        byte[] CreationDate { get; set; }

        byte[] EntryName { get; }

        string GetEntryName();

        int LeftSibling { get; set; }

        byte[] ModifyDate { get; set; }

        string Name { get; }

        ushort NameLength { get; }

        void Read(System.IO.Stream stream);

        int RightSibling { get; set; }

        void SetEntryName(string entryName);

        int SID { get; set; }

        long Size { get; set; }

        int StartSector { get; set; }

        int StateBits { get; set; }

        StgColor StgColor { get; set; }

        StgType StgType { get; set; }

        Guid StorageCLSID { get; set; }

        void Write(System.IO.Stream stream);
    }
}
