using System;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{
    /// <summary>
    ///     Used as internal template object for binary tree searches.
    /// </summary>
    internal class CFMock : CFItem
    {
        internal CFMock(string dirName, StgType dirType)
        {
            DirEntry = new DirectoryEntry(dirType);
            DirEntry.SetEntryName(dirName);
        }
    }
}