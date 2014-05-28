using System;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{
    public interface IDirectoryEntry : IComparable
    {
        int Child { get; set; }

        byte[] CreationDate { get; set; }

        byte[] EntryName { get; }

        string GetEntryName();

        int LeftSibling { get; set; }

        byte[] ModifyDate { get; set; }

        string Name { get; }

        ushort NameLength { get; set; }

        void Read(System.IO.Stream stream);

        int RightSibling { get; set; }

        void SetEntryName(string entryName);

        // ReSharper disable once InconsistentNaming
        int SID { get; set; }

        long Size { get; set; }

        int StartSetc { get; set; }

        int StateBits { get; set; }

        StgColor StgColor { get; set; }

        StgType StgType { get; set; }

        // ReSharper disable once InconsistentNaming
        Guid StorageCLSID { get; set; }

        void Write(System.IO.Stream stream);
    }
}
