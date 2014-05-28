using System;
using System.Collections.Generic;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.BinaryTree;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{

    #region Delegates
    /// <summary>
    ///     Action to apply to visited items in the OLE structured storage
    /// </summary>
    /// <param name="item">
    ///     Currently visited
    ///     <see cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFItem">item</see>
    /// </param>
    /// <example>
    ///     <code>
    ///  
    ///  //We assume that xls file should be a valid OLE compound file
    ///  const String STORAGE_NAME = "report.xls";
    ///  CompoundFile cf = new CompoundFile(STORAGE_NAME);
    /// 
    ///  FileStream output = new FileStream("LogEntries.txt", FileMode.Create);
    ///  TextWriter tw = new StreamWriter(output);
    /// 
    ///  VisitedEntryAction va = delegate(CFItem item)
    ///  {
    ///      tw.WriteLine(item.Name);
    ///  };
    /// 
    ///  cf.RootStorage.VisitEntries(va, true);
    /// 
    ///  tw.Close();
    /// 
    ///  </code>
    /// </example>
    public delegate void VisitedEntryAction(ICFItem item);

    public delegate void VisitedEntryParamsAction(ICFItem item, params object[] args);
    #endregion

    /// <summary>
    ///     Storage entity that acts like a logic container for streams or substorages in a compound file.
    /// </summary>
    public class CFStorage : CFItem, ICFStorage
    {
        #region Fields
        private BinarySearchTree<CFItem> _children;
        private NodeAction<CFItem> _internalAction;
        #endregion

        #region Constructors
        /// <summary>
        ///     Create a new CFStorage
        /// </summary>
        /// <param name="compFile">The Storage Owner - CompoundFile</param>
        internal CFStorage(CompoundFile compFile) : base(compFile)
        {
            DirEntry = new DirectoryEntry(StgType.StgStorage) {StgColor = StgColor.Black};
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
        internal BinarySearchTree<CFItem> Children
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
        /// <exception cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFDisposedException">
        ///     Raised
        ///     if trying to delete item from a closed compound file
        /// </exception>
        /// <exception cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFItemNotFound">
        ///     Raised if
        ///     item to delete is not found
        /// </exception>
        /// <example>
        ///     <code>
        ///  String filename = "report.xls";
        /// 
        ///  CompoundFile cf = new CompoundFile(filename);
        ///  CFStream foundStream = cf.RootStorage.GetStream("Workbook");
        /// 
        ///  byte[] temp = foundStream.GetData();
        /// 
        ///  Assert.IsNotNull(temp);
        /// 
        ///  cf.Close();
        ///  </code>
        /// </example>
        public ICFStream GetStream(String streamName)
        {
            CheckDisposed();

            var tmp = new CFMock(streamName, StgType.StgStream);

            CFItem outDe;

            if (Children.TryFind(tmp, out outDe) && outDe.DirEntry.StgType == StgType.StgStream)
                return outDe as CFStream;

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
        ///  String filename = "report.xls";
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
        /// <exception cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFDisposedException">
        ///     Raised
        ///     if trying to delete item from a closed compound file
        /// </exception>
        /// <exception cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFItemNotFound">
        ///     Raised if
        ///     item to delete is not found
        /// </exception>
        /// <example>
        ///     <code>
        ///  
        ///  String FILENAME = "MultipleStorage2.cfs";
        ///  CompoundFile cf = new CompoundFile(FILENAME, UpdateMode.ReadOnly, false, false);
        /// 
        ///  CFStorage st = cf.RootStorage.GetStorage("MyStorage");
        /// 
        ///  Assert.IsNotNull(st);
        ///  cf.Close();
        ///  </code>
        /// </example>
        public ICFStorage GetStorage(String storageName)
        {
            CheckDisposed();

            var tmp = new CFMock(storageName, StgType.StgStorage);

            CFItem outDe;
            if (Children.TryFind(tmp, out outDe) && outDe.DirEntry.StgType == StgType.StgStorage)
            {
                return outDe as CFStorage;
            }
            throw new CFItemNotFound("Cannot find item [" + storageName + "] within the current storage");
        }
        #endregion

        #region ExistsStorage
        /// <summary>
        ///     Checks if a child storage exists within the parent.
        /// </summary>
        /// <param name="storageName">Name of the storage to look for.</param>
        /// <returns>A boolean value indicating whether the child storage was found.</returns>
        /// <example>
        ///     <code>
        ///  String FILENAME = "MultipleStorage2.cfs";
        ///  CompoundFile cf = new CompoundFile(FILENAME, UpdateMode.ReadOnly, false, false);
        /// 
        ///  bool exists = cf.RootStorage.ExistsStorage("MyStorage");
        ///  
        ///  if exists
        ///  {
        ///      CFStorage st = cf.RootStorage.GetStorage("MyStorage");
        ///  }
        ///  
        ///  Assert.IsNotNull(st);
        ///  cf.Close();
        ///  </code>
        /// </example>
        public bool ExistsStorage(string storageName)
        {
            CheckDisposed();

            var tmp = new CFMock(storageName, StgType.StgStorage);

            CFItem outDe;
            return Children.TryFind(tmp, out outDe) && outDe.DirEntry.StgType == StgType.StgStorage;
        }
        #endregion

        #region VisitEntries
        /// <summary>
        ///     Visit all entities contained in the storage applying a user provided action
        /// </summary>
        /// <exception cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFDisposedException">
        ///     Raised
        ///     when visiting items of a closed compound file
        /// </exception>
        /// <param name="action">
        ///     User
        ///     <see cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.VisitedEntryAction">action</see>
        ///     to apply to visited entities
        /// </param>
        /// <param name="recursive">
        ///     Visiting recursion level. True means substorages are visited recursively, false indicates that
        ///     only the direct children of this storage are visited
        /// </param>
        /// <example>
        ///     <code>
        ///  const String STORAGE_NAME = "report.xls";
        ///  CompoundFile cf = new CompoundFile(STORAGE_NAME);
        /// 
        ///  FileStream output = new FileStream("LogEntries.txt", FileMode.Create);
        ///  TextWriter tw = new StreamWriter(output);
        /// 
        ///  VisitedEntryAction va = delegate(CFItem item)
        ///  {
        ///      tw.WriteLine(item.Name);
        ///  };
        /// 
        ///  cf.RootStorage.VisitEntries(va, true);
        /// 
        ///  tw.Close();
        ///  </code>
        /// </example>
        public void VisitEntries(VisitedEntryAction action, bool recursive)
        {
            CheckDisposed();

            if (action == null) return;
            var subStorages
                = new List<BinaryTreeNode<CFItem>>();

            _internalAction =
                delegate(BinaryTreeNode<CFItem> targetNode)
                {
                    action(targetNode.Value);

                    if (targetNode.Value.DirEntry.Child != DirectoryEntry.Nostream)
                        subStorages.Add(targetNode);
                };

            Children.VisitTreeInOrder(_internalAction);

            if (!recursive || subStorages.Count <= 0) return;
            foreach (var n in subStorages)
                ((CFStorage) n.Value).VisitEntries(action, true);
        }

        /// <summary>
        ///     This overload of the VisitEntries method allows the passing of a parameter arry of
        ///     objects to the delegate method.
        /// </summary>
        /// <param name="action">
        ///     User
        ///     <see cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.VisitedEntryParamsAction">action</see>
        ///     to apply to visited
        ///     entities
        /// </param>
        /// <param name="recursive">
        ///     Visiting recursion level. True means substorages are visited recursively, false
        ///     indicates that only the direct children of this storage are visited
        /// </param>
        /// <param name="args">
        ///     The arguments to pass through to the delegate method
        /// </param>
        /// <example>
        ///     <code>
        /// const String STORAGE_NAME = "report.xls";
        /// CompoundFile cf = new CompoundFile(STORAGE_NAME);
        /// 
        /// FileStream output = new FileStream("LogEntries.txt", FileMode.Create);
        /// TextWriter tw = new StreamWriter(output);
        /// 
        /// VisitedEntryParamsAction va = delegate(CFItem item, object[] args)
        /// {
        ///     var castList = (List<string />
        ///             )args[0];
        ///             castList.Add(item.Name);
        ///             };
        ///             var list = new List
        ///             <string />
        ///                 ();
        ///                 cf.RootStorage.VisitEntries(va, true, list);
        ///                 list.ForEach(tw.WriteLine);
        ///                 tw.Close();
        /// </code>
        /// </example>
        public void VisitEntries(VisitedEntryParamsAction action, bool recursive, params object[] args)
        {
            VisitedEntryAction wrappedDelegate = item => action(item, args);

            VisitEntries(wrappedDelegate, recursive);
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