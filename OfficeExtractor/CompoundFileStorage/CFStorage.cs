using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.BinaryTree;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.BinaryTree.Exceptions;
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
        /// Create a new CFStorage
        /// </summary>
        /// <param name="compFile">The Storage Owner - CompoundFile</param>
        internal CFStorage(CompoundFile compFile)
            : base(compFile)
        {
            DirEntry = new DirectoryEntry(StgType.StgStorage);
            compFile.InsertNewDirectoryEntry(DirEntry);
        }

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
                _children = CompoundFile.HasSourceStream
                    ? CompoundFile.GetChildrenTree(DirEntry.SID)
                    : new BinarySearchTree<CFItem>();

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

        #region AddStream
        /// <summary>
        /// Create a new child stream inside the current <see cref="T:OpenMcdf.CFStorage">storage</see>
        /// </summary>
        /// <param name="streamName">The new stream name</param>
        /// <returns>The new <see cref="T:OpenMcdf.CFStream">stream</see> reference</returns>
        /// <exception cref="CFDuplicatedItemException">Raised when adding an item with the same name of an existing one</exception>
        /// <exception cref="CFDisposedException">Raised when adding a stream to a closed compound file</exception>
        /// <exception cref="CFException">Raised when adding a stream with null or empty name</exception>
        /// <example>
        /// <code>
        /// 
        ///  String filename = "A_NEW_COMPOUND_FILE_YOU_CAN_WRITE_TO.cfs";
        ///
        ///  CompoundFile cf = new CompoundFile();
        ///
        ///  CFStorage st = cf.RootStorage.AddStorage("MyStorage");
        ///  CFStream sm = st.AddStream("MyStream");
        ///  byte[] b = Helpers.GetBuffer(220, 0x0A);
        ///  sm.SetData(b);
        ///
        ///  cf.Save(filename);
        ///  
        /// </code>
        /// </example>
        public CFStream AddStream(string streamName)
        {
            CheckDisposed();

            if (string.IsNullOrEmpty(streamName))
                throw new CFException("Stream name cannot be null or empty");


            // Add new Stream directory entry
            var stream = new CFStream(CompoundFile);
            stream.DirEntry.SetEntryName(streamName);

            try
            {
                // Add object to Siblings tree
                Children.Add(stream);

                //Rethread children tree...
                CompoundFile.RefreshIterative(Children.Root);

                // ... and set the root of the tree as new child of the current item directory entry
                DirEntry.Child = Children.Root.Value.DirEntry.SID;
            }
            catch (BSTDuplicatedException)
            {
                CompoundFile.ResetDirectoryEntry(stream.DirEntry.SID);
                throw new CFDuplicatedItemException("An entry with name '" + streamName +
                                                    "' is already present in storage '" + Name + "' ");
            }

            return stream;
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

        #region AddStorage
        /// <summary>
        /// Create new child storage directory inside the current storage.
        /// </summary>
        /// <param name="storageName">The new storage name</param>
        /// <returns>Reference to the new <see cref="T:OpenMcdf.CFStorage">storage</see></returns>
        /// <exception cref="T:OpenMcdf.CFDuplicatedItemException">Raised when adding an item with the same name of an existing one</exception>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised when adding a storage to a closed compound file</exception>
        /// <exception cref="T:OpenMcdf.CFException">Raised when adding a storage with null or empty name</exception>
        /// <example>
        /// <code>
        /// 
        ///  String filename = "A_NEW_COMPOUND_FILE_YOU_CAN_WRITE_TO.cfs";
        ///
        ///  CompoundFile cf = new CompoundFile();
        ///
        ///  CFStorage st = cf.RootStorage.AddStorage("MyStorage");
        ///  CFStream sm = st.AddStream("MyStream");
        ///  byte[] b = Helpers.GetBuffer(220, 0x0A);
        ///  sm.SetData(b);
        ///
        ///  cf.Save(filename);
        ///  
        /// </code>
        /// </example>
        public CFStorage AddStorage(string storageName)
        {
            CheckDisposed();

            if (string.IsNullOrEmpty(storageName))
                throw new CFException("Stream name cannot be null or empty");

            // Add new Storage directory entry
            var storage = new CFStorage(CompoundFile);
            storage.DirEntry.SetEntryName(storageName);

            try
            {
                // Add object to Siblings tree
                Children.Add(storage);
            }
            catch (BSTDuplicatedException)
            {

                CompoundFile.ResetDirectoryEntry(storage.DirEntry.SID);
                throw new CFDuplicatedItemException("An entry with name '" + storageName +
                                                    "' is already present in storage '" + Name + "' ");
            }


            CompoundFile.RefreshIterative(Children.Root);
            DirEntry.Child = Children.Root.Value.DirEntry.SID;
            return storage;
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
    }
}