using System;
using System.Collections.Generic;
using System.IO;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.BinaryTree;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Interfaces;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{
    #region Enum CFSVersion
    /// <summary>
    ///     Binary File Format Version. Sector size  is 512 byte for version 3, 4096 for version 4
    /// </summary>
    public enum CFSVersion
    {
        /// <summary>
        ///     Compound file version 3 - The default and most common version available. Sector size 512 bytes, 2GB max file size.
        /// </summary>
        Ver3 = 3,

        /// <summary>
        ///     Compound file version 4 - Sector size is 4096 bytes. Using this version could bring some compatibility problem with
        ///     existing applications.
        /// </summary>
        Ver4 = 4
    }
    #endregion

    /// <summary>
    ///     Standard Microsoft's; Compound File implementation. It is also known as OLE/COM structured storage
    ///     and contains a hierarchy of storage and stream objects providing efficent storage of multiple kinds 
    ///     of documents in a single file. Version 3 and 4 of specifications are supported.
    /// </summary>
    public class CompoundFile : IDisposable, ICompoundFile
    {
        #region Fields
        private List<IDirectoryEntry> _directoryEntries = new List<IDirectoryEntry>();
        internal int LockSectorId = -1;
        internal bool TransactionLockAdded = false;
        internal bool TransactionLockAllocated = false;

        /// <summary>
        ///     CompoundFile header
        /// </summary>
        private Header _header;

        private object _lockObject = new Object();
        private CFStorage _rootStorage;

        /// <summary>
        /// Contains all the sectors in the compound file storage
        /// </summary>
        private SectorCollection _sectors = new SectorCollection();

        /// <summary>
        ///     Compound underlying stream. Null when new CF has been created.
        /// </summary>
        internal Stream SourceStream = null;
        #endregion

        #region Properties
        internal bool IsClosed { get; private set; }

        /// <summary>
        ///     Gets the root entry of the compound file
        /// </summary>
        internal IDirectoryEntry RootEntry
        {
            get { return _directoryEntries[0]; }
        }

        /// <summary>
        ///     Return true if this compound file has been
        ///     loaded from an existing file or stream
        /// </summary>
        public bool HasSourceStream
        {
            get { return SourceStream != null; }
        }

        /// <summary>
        ///     The entry point object that represents the
        ///     root of the structures tree to get or set storage or
        ///     stream data.
        /// </summary>
        /// <example>
        ///     <code>
        ///  
        ///     //Create a compound file
        ///     string FILENAME = "MyFileName.cfs";
        ///     CompoundFile ncf = new CompoundFile();
        /// 
        ///     CFStorage l1 = ncf.RootStorage.AddStorage("Storage Level 1");
        /// 
        ///     l1.AddStream("l1ns1");
        ///     l1.AddStream("l1ns2");
        ///     l1.AddStream("l1ns3");
        ///     CFStorage l2 = l1.AddStorage("Storage Level 2");
        ///     l2.AddStream("l2ns1");
        ///     l2.AddStream("l2ns2");
        /// 
        ///     ncf.Save(FILENAME);
        ///     ncf.Close();
        ///  </code>
        /// </example>
        public ICFStorage RootStorage
        {
            get { return _rootStorage; }
        }

        public CFSVersion Version
        {
            get { return (CFSVersion)_header.MajorVersion; }
        }
        #endregion

        #region Constructors
        /// <summary>
        ///     Load an existing compound file.
        /// </summary>
        /// <param name="fileName">Compound file to read from</param>
        /// <example>
        ///     <code>
        ///  //A xls file should have a Workbook stream
        ///  string filename = "report.xls";
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
        /// <remarks>
        ///     File will be open in read-only mode. Version will be automatically recognized
        ///     from the file. Sector recycle is turned off to achieve the best reading
        ///     performance in most common scenarios.
        /// </remarks>
        public CompoundFile(string fileName)
        {
            FileStream fileStream = null;

            try
            {
                fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
                LoadStream(fileStream);
            }
            catch (Exception)
            {
                if (fileStream != null)
                    fileStream.Close();

                throw;
            }
        }

        /// <summary>
        ///     Load an existing compound file from a stream.
        /// </summary>
        /// <param name="stream">Streamed compound file</param>
        /// <example>
        ///     <code>
        ///  
        ///  string filename = "reportREAD.xls";
        ///    
        ///  FileStream fs = new FileStream(filename, FileMode.Open);
        ///  CompoundFile cf = new CompoundFile(fs);
        ///  CFStream foundStream = cf.RootStorage.GetStream("Workbook");
        /// 
        ///  byte[] temp = foundStream.GetData();
        /// 
        ///  Assert.IsNotNull(temp);
        /// 
        ///  cf.Close();
        /// 
        ///  </code>
        /// </example>
        /// <exception cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFException">
        ///     Raised when
        ///     trying to open a non-seekable stream
        /// </exception>
        /// <exception cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFException">
        ///     Raised
        ///     stream is null
        /// </exception>
        public CompoundFile(Stream stream)
        {
            LoadStream(stream);
        }
        #endregion
        
        #region LoadStream
        /// <summary>
        ///     Load compound file from an existing stream.
        /// </summary>
        /// <param name="stream">Stream to load compound file from</param>
        private void LoadStream(Stream stream)
        {
            if (stream == null)
                throw new CFException("Stream parameter cannot be null");

            if (!stream.CanSeek)
                throw new CFException("Cannot load a non-seekable Stream");

            stream.Seek(0, SeekOrigin.Begin);

            try
            {
                _header = new Header();
                _directoryEntries = new List<IDirectoryEntry>();

                SourceStream = stream;

                _header.Read(stream);

                var numberOfSectors = Ceiling(((stream.Length - GetSectorSize())/(double) GetSectorSize()));

                if (stream.Length > 0x7FFFFF0)
                    TransactionLockAllocated = true;

                _sectors = new SectorCollection();

                for (var i = 0; i < numberOfSectors; i++)
                    _sectors.Add(null);

                LoadDirectories();

                _rootStorage
                    = new CFStorage(this, _directoryEntries[0]);
            }
            catch (Exception)
            {
                stream.Close();
                throw;
            }
        }
        #endregion

        #region GetAllNamedEntries
        /// <summary>
        ///     Returns a list of all entries with part of the name <see cref="entryName"/>
        /// </summary>
        /// <param name="entryName">Name of entries to retrieve</param>
        /// <returns>A list of name-matching entries</returns>
        /// <remarks>
        ///     This function is aimed to speed up entity lookup in
        ///     flat-structure files (only one or little more known entries)
        ///     without the performance penalty related to entities hierarchy constraints.
        ///     There is no implied hierarchy in the returned list.
        /// </remarks>
        public IList<ICFItem> GetAllNamedEntries(string entryName)
        {
            var result = new List<ICFItem>();

            foreach (var directoryEntry in _directoryEntries)
            {
                if (directoryEntry.Name.ToUpperInvariant().Contains(entryName.ToUpperInvariant()) &&
                    directoryEntry.StgType != StgType.StgInvalid)
                    result.Add(directoryEntry.StgType == StgType.StgStorage
                        ? new CFStorage(this, directoryEntry)
                        : (ICFItem) new CFStream(this, directoryEntry));
            }

            return result;
        }
        #endregion

        #region Close
        /// <summary>
        ///     Close the Compound File object <see cref="T:OpenMcdf.CompoundFile">CompoundFile</see> and
        ///     free all associated resources (e.g. open file handle and allocated memory).
        ///     <remarks>
        ///         When the <see cref="T:OpenMcdf.CompoundFile.Close()">Close</see> method is called,
        ///         all the associated stream and storage objects are invalidated:
        ///         any operation invoked on them will produce a
        ///         <see cref="T:OpenMcdf.CFDisposedException">CFDisposedException</see>.
        ///     </remarks>
        /// </summary>
        /// <example>
        ///     <code>
        ///     const string FILENAME = "CompoundFile.cfs";
        ///     CompoundFile cf = new CompoundFile(FILENAME);
        /// 
        ///     CFStorage st = cf.RootStorage.GetStorage("MyStorage");
        ///     cf.Close();
        /// 
        ///     try
        ///     {
        ///         byte[] temp = st.GetStream("MyStream").GetData();
        ///         
        ///         // The following line will fail because back-end object has been closed
        ///         Assert.Fail("Stream without media");
        ///     }
        ///     catch (Exception ex)
        ///     {
        ///         Assert.IsTrue(ex is CFDisposedException);
        ///     }
        ///  </code>
        /// </example>
        public void Close()
        {
            ((IDisposable) this).Dispose();
        }
        #endregion

        #region GetSectorSize
        /// <summary>
        ///     Returns the size of standard sectors switching on CFS version (3 or 4)
        /// </summary>
        /// <returns>Standard sector size</returns>
        internal int GetSectorSize()
        {
            return 2 << (_header.SectorShift - 1);
        }
        #endregion

        #region GetDIFATSectorChain
        /// <summary>
        ///     Gets the DIFAT Sector chain
        /// </summary>
        /// <returns>A list of DIFAT sectors</returns>
        /// <exception cref="CFCorruptedFileException">Raised when a compound file is corrupt</exception>
        private List<Sector> GetDIFATSectorChain()
        {
            var result = new List<Sector>();

            if (_header.DIFATSectorsNumber == 0) return result;
            var validationCount = (int) _header.DIFATSectorsNumber;

            var sector = _sectors[_header.FirstDIFATSectorId];

            if (sector == null) //Lazy loading
            {
                sector = new Sector(GetSectorSize(), SourceStream)
                {
                    Type = SectorType.DIFAT,
                    Id = _header.FirstDIFATSectorId
                };
                _sectors[_header.FirstDIFATSectorId] = sector;
            }

            result.Add(sector);

            while (validationCount >= 0)
            {
                var nextSectorId = BitConverter.ToInt32(sector.GetData(), GetSectorSize() - 4);

                // Strictly speaking, the following condition is not correct from a specification point of view:
                // only ENDOFCHAIN should break DIFAT chain but 
                // a lot of existing compound files use FREESECT as DIFAT chain termination
                if (nextSectorId == Sector.FreeSector || nextSectorId == Sector.Endofchain) break;

                validationCount--;

                if (validationCount < 0)
                {
                    Close();
                    throw new CFCorruptedFileException("DIFAT sectors count mismatched. Corrupted compound file");
                }

                sector = _sectors[nextSectorId];

                if (sector == null)
                {
                    sector = new Sector(GetSectorSize(), SourceStream) {Id = nextSectorId};
                    _sectors[nextSectorId] = sector;
                }

                result.Add(sector);
            }

            return result;
        }
        #endregion

        #region GetFATSectorChain
        /// <summary>
        ///     Gets the FAT sector chain
        /// </summary>
        /// <returns>List of FAT sectors</returns>
        private List<Sector> GetFATSectorChain()
        {
            const int numberHeaderFATEntry = 109; //Number of FAT sectors id in the header

            var result = new List<Sector>();
            int nextSectorId;
            var difatSectors = GetDIFATSectorChain();
            var idx = 0;

            // Read FAT entries from the header FAT entry array (max 109 entries)
            while (idx < _header.FATSectorsNumber && idx < numberHeaderFATEntry)
            {
                nextSectorId = _header.DIFAT[idx];
                var sector = _sectors[nextSectorId];

                if (sector == null)
                {
                    sector = new Sector(GetSectorSize(), SourceStream) {Id = nextSectorId, Type = SectorType.FAT};
                    _sectors[nextSectorId] = sector;
                }

                result.Add(sector);

                idx++;
            }

            // Is there any DIFAT sector containing other FAT entries ?
            if (difatSectors.Count <= 0) return result;
            var difatStream
                = new StreamViewer
                    (
                    difatSectors,
                    GetSectorSize(),
                    _header.FATSectorsNumber > numberHeaderFATEntry
                        ? (_header.FATSectorsNumber - numberHeaderFATEntry)*4
                        : 0,
                    SourceStream
                    );

            var nextDIFATSectorBuffer = new byte[4];

            difatStream.Read(nextDIFATSectorBuffer, 0, 4);
            nextSectorId = BitConverter.ToInt32(nextDIFATSectorBuffer, 0);

            var i = 0;
            var nFat = numberHeaderFATEntry;

            while (nFat < _header.FATSectorsNumber)
            {
                if (difatStream.Position == ((GetSectorSize() - 4) + i*GetSectorSize()))
                {
                    difatStream.Seek(4, SeekOrigin.Current);
                    i++;
                    continue;
                }

                var sector = _sectors[nextSectorId];

                if (sector == null)
                {
                    sector = new Sector(GetSectorSize(), SourceStream) {Type = SectorType.FAT, Id = nextSectorId};
                    _sectors[nextSectorId] = sector;
                }

                result.Add(sector);

                difatStream.Read(nextDIFATSectorBuffer, 0, 4);
                nextSectorId = BitConverter.ToInt32(nextDIFATSectorBuffer, 0);
                nFat++;
            }

            return result;
        }
        #endregion

        #region GetNormalSectorChain
        /// <summary>
        ///     Gets a standard sector chain
        /// </summary>
        /// <param name="sectorId">First SecID of the required chain</param>
        /// <returns>A list of sectors</returns>
        /// <exception cref="CFCorruptedFileException">Raised when the compound file is corrupt</exception>
        private List<Sector> GetNormalSectorChain(int sectorId)
        {
            var result
                = new List<Sector>();

            var nextSectorId = sectorId;

            var fatSectors = GetFATSectorChain();

            var fatStream
                = new StreamViewer(fatSectors, GetSectorSize(), fatSectors.Count*GetSectorSize(), SourceStream);

            while (true)
            {
                if (nextSectorId == Sector.Endofchain) break;

                if (nextSectorId >= _sectors.Count)
                    throw new CFCorruptedFileException(
                        string.Format(
                            "Next Sector ID reference an out of range sector. NextID : {0} while sector count {1}",
                            nextSectorId, _sectors.Count));

                var sector = _sectors[nextSectorId];
                if (sector == null)
                {
                    sector = new Sector(GetSectorSize(), SourceStream) {Id = nextSectorId, Type = SectorType.Normal};
                    _sectors[nextSectorId] = sector;
                }

                result.Add(sector);

                fatStream.Seek(nextSectorId*4, SeekOrigin.Begin);
                var next = fatStream.ReadInt32();

                if (next != nextSectorId)
                    nextSectorId = next;
                else
                    throw new CFCorruptedFileException("Cyclic sector chain found. File is corrupted");
            }

            return result;
        }
        #endregion

        #region GetMiniSectorChain
        /// <summary>
        ///     Gets a mini sector chain
        /// </summary>
        /// <param name="sectorId">First SecID of the required chain</param>
        /// <returns>A list of mini sectors (64 bytes)</returns>
        private List<Sector> GetMiniSectorChain(int sectorId)
        {
            var result = new List<Sector>();

            if (sectorId == Sector.Endofchain) return result;

            var miniFAT = GetNormalSectorChain(_header.FirstMiniFATSectorId);
            var miniStream = GetNormalSectorChain(RootEntry.StartSector);

            var miniFATView = new StreamViewer(miniFAT, GetSectorSize(),
                _header.MiniFATSectorsNumber*Sector.MinisectorSize,
                SourceStream);

            var miniStreamView = new StreamViewer(miniStream, GetSectorSize(), _rootStorage.Size, SourceStream);

            var miniFATReader = new BinaryReader(miniFATView);

            var nextSectorId = sectorId;

            while (true)
            {
                if (nextSectorId == Sector.Endofchain)
                    break;

                var miniSector = new Sector(Sector.MinisectorSize, SourceStream) {Id = nextSectorId, Type = SectorType.Mini};

                miniStreamView.Seek(nextSectorId*Sector.MinisectorSize, SeekOrigin.Begin);
                miniStreamView.Read(miniSector.GetData(), 0, Sector.MinisectorSize);

                result.Add(miniSector);

                miniFATView.Seek(nextSectorId*4, SeekOrigin.Begin);
                nextSectorId = miniFATReader.ReadInt32();
            }
            return result;
        }
        #endregion

        #region GetSectorChain
        /// <summary>
        ///     Gets a sector chain from a compound file given the first sector ID and the required sector type.
        /// </summary>
        /// <param name="sectorId">First chain sector's id </param>
        /// <param name="chainType">Type of Sectors in the required chain (mini sectors, normal sectors or FAT)</param>
        /// <returns>A list of Sectors as the result of their concatenation</returns>
        /// <exception cref="CFException">Raised when an unsuported chain type is given</exception>
        internal List<Sector> GetSectorChain(int sectorId, SectorType chainType)
        {
            switch (chainType)
            {
                case SectorType.DIFAT:
                    return GetDIFATSectorChain();

                case SectorType.FAT:
                    return GetFATSectorChain();

                case SectorType.Normal:
                    return GetNormalSectorChain(sectorId);

                case SectorType.Mini:
                    return GetMiniSectorChain(sectorId);

                default:
                    throw new CFException("Unsupported chain type");
            }
        }
        #endregion

        #region InsertNewDirectoryEntry
        /// <summary>
        /// Inserts a new directory entry
        /// </summary>
        /// <param name="directoryEntry"></param>
        internal void InsertNewDirectoryEntry(IDirectoryEntry directoryEntry)
        {
            // If we are not adding an invalid dirEntry as
            // in a normal loading from file (invalid dirs MAY pad a sector)
            if (directoryEntry != null)
            {
                // Find first available invalid slot (if any) to reuse it
                for (var i = 0; i < _directoryEntries.Count; i++)
                {
                    if (_directoryEntries[i].StgType != StgType.StgInvalid) continue;
                    _directoryEntries[i] = directoryEntry;
                    directoryEntry.SID = i;
                    return;
                }
            }

            // No invalid directory entry found
            _directoryEntries.Add(directoryEntry);
            if (directoryEntry != null) directoryEntry.SID = _directoryEntries.Count - 1;
        }
        #endregion

        #region ResetDirectoryEntry
        /// <summary>
        ///     Reset a directory entry setting it to StgInvalid in the Directory.
        /// </summary>
        /// <param name="sid">Sid of the directory to invalidate</param>
        internal void ResetDirectoryEntry(int sid)
        {
            _directoryEntries[sid] = new DirectoryEntry(StgType.StgInvalid);
        }
        #endregion

        #region GetChildrenTree
        /// <summary>
        /// Returns all the children of the given <see cref="sid"/> sibling
        /// </summary>
        /// <param name="sid"></param>
        /// <returns></returns>
        internal BinarySearchTree<CFItem> GetChildrenTree(int sid)
        {
            var binarySearchTree = new BinarySearchTree<CFItem>(new CFItemComparer());

            // Load children from their original tree.
            LoadChildren(binarySearchTree, _directoryEntries[sid]);

            // Rebuild of (Red)-Black tree of entry children.
            binarySearchTree.VisitTreeInOrder(RefreshSIDs);

            return binarySearchTree;
        }
        #endregion

        #region LoadChildren
        /// <summary>
        /// Loads all the children of the sibling
        /// </summary>
        /// <param name="binarySearchTree"></param>
        /// <param name="directoryEntry"></param>
        private void LoadChildren(ICollection<CFItem> binarySearchTree, IDirectoryEntry directoryEntry)
        {
            if (directoryEntry.Child == DirectoryEntry.Nostream) return;
            if (_directoryEntries[directoryEntry.Child].StgType == StgType.StgInvalid) return;

            if (_directoryEntries[directoryEntry.Child].StgType == StgType.StgStream)
                binarySearchTree.Add(new CFStream(this, _directoryEntries[directoryEntry.Child]));
            else
                binarySearchTree.Add(new CFStorage(this, _directoryEntries[directoryEntry.Child]));

            LoadSiblings(_directoryEntries[directoryEntry.Child]);
        }
        #endregion
        
        #region LoadSiblings
        /// <summary>
        /// Doubling methods allows iterative behavior while avoiding to insert duplicate items
        /// </summary>
        /// <param name="directoryEntry"></param>
        private void LoadSiblings(IDirectoryEntry directoryEntry)
        {
            while (true)
            {
                if (directoryEntry.LeftSibling != DirectoryEntry.Nostream)
                {
                    // If there're more left siblings load them...
                    LoadSiblings(_directoryEntries[directoryEntry.LeftSibling]);
                }

                if (directoryEntry.RightSibling != DirectoryEntry.Nostream)
                {
                    // If there're more right siblings load them...
                    directoryEntry = _directoryEntries[directoryEntry.RightSibling];
                    continue;
                }
                break;
            }
        }
        #endregion

        #region LoadDirectories
        /// <summary>
        ///     Load directory entries from compound file. Header and FAT MUST be already loaded.
        /// </summary>
        private void LoadDirectories()
        {
            var directoryChain = GetSectorChain(_header.FirstDirectorySectorId, SectorType.Normal);

            if (_header.FirstDirectorySectorId == Sector.Endofchain)
                _header.FirstDirectorySectorId = directoryChain[0].Id;

            var directoryReader = new StreamViewer(directoryChain, GetSectorSize(), directoryChain.Count*GetSectorSize(),
                SourceStream);


            while (directoryReader.Position < directoryChain.Count*GetSectorSize())
            {
                var directoryEntry = new DirectoryEntry(StgType.StgInvalid);

                //We are not inserting dirs. Do not use 'InsertNewDirectoryEntry'
                directoryEntry.Read(directoryReader);
                _directoryEntries.Add(directoryEntry);
                directoryEntry.SID = _directoryEntries.Count - 1;
            }
        }
        #endregion

        #region RefreshSIDs
        /// <summary>
        /// Refreshes all the SIDS
        /// </summary>
        /// <param name="node"></param>
        internal void RefreshSIDs(BinaryTreeNode<CFItem> node)
        {
            if (node.Value == null) return;
            if (node.Left != null && (node.Left.Value.DirEntry.StgType != StgType.StgInvalid))
                node.Value.DirEntry.LeftSibling = node.Left.Value.DirEntry.SID;
            else
                node.Value.DirEntry.LeftSibling = DirectoryEntry.Nostream;

            if (node.Right != null && (node.Right.Value.DirEntry.StgType != StgType.StgInvalid))
                node.Value.DirEntry.RightSibling = node.Right.Value.DirEntry.SID;
            else
                node.Value.DirEntry.RightSibling = DirectoryEntry.Nostream;
        }
        #endregion

        #region RefreshIterative
        internal void RefreshIterative(BinaryTreeNode<CFItem> node)
        {
            while (true)
            {
                if (node == null) return;
                RefreshSIDs(node);
                RefreshIterative(node.Left);
                node = node.Right;
            }
        }
        #endregion

        #region GetData
        /// <summary>
        /// Returns data from the <see cref="cFStream"/> from the <see cref="offset"/>
        /// </summary>
        /// <param name="cFStream"></param>
        /// <param name="offset"></param>
        /// <param name="count"></param>
        /// <exception cref="CFDisposedException">Raised when the file is already disposed</exception>
        /// <returns></returns>
        internal byte[] GetData(CFStream cFStream, long offset, ref int count)
        {
            if (IsClosed)
                throw new CFDisposedException("Compound File closed: cannot access data");

            var directoryEntry = cFStream.DirEntry;

            count = (int) Math.Min(directoryEntry.Size - offset, count);

            StreamViewer streamViewer;

            if (directoryEntry.Size < _header.MinSizeStandardStream)
                streamViewer
                    = new StreamViewer(GetSectorChain(directoryEntry.StartSector, SectorType.Mini), Sector.MinisectorSize, directoryEntry.Size,
                        SourceStream);
            else
                streamViewer = new StreamViewer(GetSectorChain(directoryEntry.StartSector, SectorType.Normal), GetSectorSize(), directoryEntry.Size,
                    SourceStream);

            var result = new byte[count];

            streamViewer.Seek(offset, SeekOrigin.Begin);
            streamViewer.Read(result, 0, result.Length);

            return result;
        }

        /// <summary>
        /// Returns data from the <see cref="cFStream"/>
        /// </summary>
        /// <param name="cFStream"></param>
        /// <exception cref="CFDisposedException">Raised when the file is already disposed</exception>
        /// <returns></returns>
        internal byte[] GetData(CFStream cFStream)
        {
            if (IsClosed)
                throw new CFDisposedException("Compound File closed: cannot access data");

            byte[] result;

            var directoryEntry = cFStream.DirEntry;

            //IDirectoryEntry root = directoryEntries[0];

            if (directoryEntry.Size < _header.MinSizeStandardStream)
            {
                var miniStreamViewer
                    = new StreamViewer(GetSectorChain(directoryEntry.StartSector, SectorType.Mini), Sector.MinisectorSize, directoryEntry.Size,
                        SourceStream);

                var br = new BinaryReader(miniStreamViewer);

                result = br.ReadBytes((int) directoryEntry.Size);
                br.Close();
            }
            else
            {
                var streamViewer
                    = new StreamViewer(GetSectorChain(directoryEntry.StartSector, SectorType.Normal), GetSectorSize(), directoryEntry.Size,
                        SourceStream);

                result = new byte[(int) directoryEntry.Size];

                streamViewer.Read(result, 0, result.Length);
            }

            return result;
        }
        #endregion

        #region Ceiling
        private static int Ceiling(double d)
        {
            return (int) Math.Ceiling(d);
        }
        #endregion

        #region Close
        internal void Close(bool close)
        {
            ((IDisposable) this).Dispose();
        }
        #endregion
        
        #region IDisposable Members
        /// <summary>
        ///     When called from user code, release all resources, otherwise, in the case runtime called it,
        ///     only unmanagd resources are released.
        /// </summary>
        /// <param name="disposing">If true, method has been called from User code, if false it's been called from .net runtime</param>
        protected virtual void Dispose(bool disposing)
        {
            try
            {
                if (!IsClosed)
                {
                    lock (_lockObject)
                    {
                        if (disposing)
                        {
                            // Call from user code...

                            if (_sectors != null)
                            {
                                _sectors.Clear();
                                _sectors = null;
                            }

                            _rootStorage = null; // Some problem releasing resources...
                            _header = null;
                            _directoryEntries.Clear();
                            _directoryEntries = null;
                            _lockObject = null;
                        }

                        if (SourceStream != null)
                            SourceStream.Close();
                    }
                }
            }
            finally
            {
                IsClosed = true;
            }
        }

        void IDisposable.Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}