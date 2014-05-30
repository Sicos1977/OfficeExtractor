using System.Collections.Generic;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.BinaryTree;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Interfaces;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{
    /// <summary>
    ///     Storage entity that acts like a logic container for streams or substorages in a compound file.
    /// </summary>
    public class CFStorage : CFItem, ICFStorage
    {
        #region Fields
        private BinarySearchTree<CFItem> _children;
        #endregion

        #region Constructor
        /// <summary>
        ///     Create a CFStorage using an existing directory (previously loaded).
        /// </summary>
        /// <param name="compFile">The Storage Owner - CompoundFile</param>
        /// <param name="dirEntry">An existing Directory Entry</param>
        internal CFStorage(CompoundFile compFile, IDirectoryEntry dirEntry) : base(compFile)
        {
            if (dirEntry == null || dirEntry.SID < 0)
                throw new CFException("Attempting to create a CFStorage using an unitialized directory");

            DirEntry = dirEntry;
        }
        #endregion

        #region Children
        public BinarySearchTree<CFItem> Children
        {
            get
            {
                // Lazy loading of children tree.
                if (_children != null) return _children;
                _children = CompoundFile.HasSourceStream ? LoadChildren(DirEntry.SID) : new BinarySearchTree<CFItem>();

                return _children;
            }
        }
        #endregion

        #region GetStream
        /// <summary>
        ///     Get a named
        ///     <see cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFStream">stream</see>
        ///     contained in the current storage if existing.
        /// </summary>
        /// <param name="streamName">Name of the stream to look for</param>
        /// <returns>A stream reference if existing</returns>
        /// <exception cref="CFItemNotFound">Raised if <see cref="streamName"/> is not found</exception>
        public ICFStream GetStream(string streamName)
        {
            CheckDisposed();

            var cfMock = new CFMock(streamName, StgType.StgStream);

            CFItem directoryEntry;

            if (Children.TryFind(cfMock, out directoryEntry) && directoryEntry.DirEntry.StgType == StgType.StgStream)
                return directoryEntry as CFStream;

            throw new CFItemNotFound("Cannot find item [" + streamName + "] within the current storage");
        }
        #endregion

        #region ExistsStream
        /// <summary>
        ///     Checks whether a child stream exists in the parent.
        /// </summary>
        /// <param name="streamName">Name of the stream to look for</param>
        /// <returns>A boolean value indicating whether the child stream exists.</returns>
        /// <example>
        ///     <code>
        ///  string filename = "report.xls";
        /// 
        ///  CompoundFile cf = new CompoundFile(filename);
        ///  
        ///  bool exists = ExistsStream("Workbook");
        ///  
        ///  if exists
        ///  {
        ///      CFStream foundStream = cf.RootStorage.GetStream("Workbook");
        ///  
        ///      byte[] temp = foundStream.GetData();
        ///  }
        /// 
        ///  Assert.IsNotNull(temp);
        /// 
        ///  cf.Close();
        ///  </code>
        /// </example>
        public bool ExistsStream(string streamName)
        {
            CheckDisposed();

            var tmp = new CFMock(streamName, StgType.StgStream);

            CFItem outDe;
            return Children.TryFind(tmp, out outDe) && outDe.DirEntry.StgType == StgType.StgStream;
        }
        #endregion

        #region GetStorage
        /// <summary>
        ///     Get a named storage contained in the current one if existing.
        /// </summary>
        /// <param name="storageName">Name of the storage to look for</param>
        /// <returns>A storage reference if existing.</returns>
        /// <exception cref="CFItemNotFound">Raised if <see cref="storageName"/> is not found</exception>
        public ICFStorage GetStorage(string storageName)
        {
            CheckDisposed();

            var cfMock = new CFMock(storageName, StgType.StgStorage);

            CFItem directoryEntry;
            if (Children.TryFind(cfMock, out directoryEntry) && directoryEntry.DirEntry.StgType == StgType.StgStorage)
                return directoryEntry as CFStorage;
            
            throw new CFItemNotFound("Cannot find item [" + storageName + "] within the current storage");
        }
        #endregion

        #region ExistsStorage
        /// <summary>
        ///     Checks if a child storage exists within the parent.
        /// </summary>
        /// <param name="storageName">Name of the storage to look for.</param>
        /// <returns>A boolean value indicating whether the child storage was found.</returns>
        public bool ExistsStorage(string storageName)
        {
            CheckDisposed();

            var cfMock = new CFMock(storageName, StgType.StgStorage);

            CFItem directoryEntry;
            return Children.TryFind(cfMock, out directoryEntry) && directoryEntry.DirEntry.StgType == StgType.StgStorage;
        }
        #endregion
        
        #region LoadChildren
        private BinarySearchTree<CFItem> LoadChildren(int sid)
        {
            return CompoundFile.GetChildrenTree(sid);
        }
        #endregion
    }
}