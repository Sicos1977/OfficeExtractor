using System;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Interfaces;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{
    /// <summary>
    ///     OLE structured storage
    ///     <see cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFStream">stream</see>
    ///     Object
    ///     It is contained inside a Storage object in a file-directory relationship and indexed by its name.
    /// </summary>
    public class CFStream : CFItem, ICFStream
    {
        #region Constructors
        internal CFStream(CompoundFile sectorManager) : base(sectorManager)
        {
            DirEntry = new DirectoryEntry(StgType.StgStream);
            sectorManager.InsertNewDirectoryEntry(DirEntry);
        }

        internal CFStream(CompoundFile sectorManager, IDirectoryEntry dirEntry) : base(sectorManager)
        {
            if (dirEntry == null || dirEntry.SID < 0)
                throw new CFException("Attempting to add a CFStream using an unitialized directory");

            DirEntry = dirEntry;
        }
        #endregion

        #region GetData
        /// <summary>
        ///     Get the data associated with the stream object.
        /// </summary>
        /// <example>
        ///     <code>
        ///     CompoundFile cf2 = new CompoundFile("AFileName.cfs");
        ///     CFStream st = cf2.RootStorage.GetStream("MyStream");
        ///     byte[] buffer = st.GetData();
        /// </code>
        /// </example>
        /// <returns>Array of byte containing stream data</returns>
        /// <exception cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFDisposedException">
        ///     Raised when the owner compound file has been closed.
        /// </exception>
        public Byte[] GetData()
        {
            CheckDisposed();

            return CompoundFile.GetData(this);
        }

        /// <summary>
        ///     Get <paramref name="count" /> bytes associated with the stream object, starting from
        ///     a provided <paramref name="offset" />. When method returns, count will contain the
        ///     effective count of bytes read.
        /// </summary>
        /// <example>
        ///     <code>
        /// CompoundFile cf = new CompoundFile("AFileName.cfs");
        /// CFStream st = cf.RootStorage.GetStream("MyStream");
        /// int count = 8;
        /// // The stream is supposed to have a length greater than offset + count
        /// byte[] data = st.GetData(20, ref count);  
        /// cf.Close();
        /// </code>
        /// </example>
        /// <returns>Array of byte containing stream data</returns>
        /// <exception cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFDisposedException">
        ///     Raised when the owner compound file has been closed.
        /// </exception>
        public Byte[] GetData(long offset, ref int count)
        {
            CheckDisposed();

            return CompoundFile.GetData(this, offset, ref count);
        }
        #endregion
    }
}