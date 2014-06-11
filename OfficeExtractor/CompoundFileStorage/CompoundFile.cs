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
    ///     Binary File Format Version. Sector size  is 512 byte for version 3,
    ///     4096 for version 4
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

    #region UpdateMode
    /// <summary>
    ///     Update mode of the compound file.
    ///     Default is ReadOnly.
    /// </summary>
    public enum UpdateMode
    {
        /// <summary>
        ///     ReadOnly update mode prevents overwriting
        ///     of the opened file.
        ///     Data changes are allowed but they have to be
        ///     persisted on a different file when required.
        /// </summary>
        ReadOnly,

        /// <summary>
        ///     Update mode allows subsequent data changing operations
        ///     to be persisted directly on the opened file or stream
        ///     using the
        ///     <see cref="M:DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.CompoundFile.Commit">Commit</see>
        ///     method when required. Warning: this option may cause existing data loss if misused.
        /// </summary>
        Update
    }
    #endregion

    /// <summary>
    ///     Standard Microsoft; Compound File implementation. It is also known as OLE/COM structured storage
    ///     and contains a hierarchy of storage and stream objects providing  efficent storage of multiple
    ///     kinds of documents in a single file. Version 3 and 4 of specifications are supported.
    /// </summary>
    public class CompoundFile : IDisposable
    {
        #region Consts
        /// <summary>
        ///     Initial capacity of the flushing queue used
        ///     to optimize commit writing operations
        /// </summary>
        private const int FlushingQueueSize = 6000;

        /// <summary>
        ///     Maximum size of the flushing buffer used
        ///     to optimize commit writing operations
        /// </summary>
        private const int FlushingBufferMaxSize = 1024*1024*16;

        /// <summary>
        ///     Number of DIFAT entries in the header
        /// </summary>
        private const int HeaderDIFATEntriesCount = 109;
        #endregion

        #region Fields
        /// <summary>
        ///     Sector ID Size (int)
        /// </summary>
        private const int SizeOfSID = 4;

        /// <summary>
        ///     Number of FAT entries in a DIFAT Sector
        /// </summary>
        private readonly int _difatSectorFATEntriesCount = 127;

        /// <summary>
        ///     Flag for unallocated sector zeroing out.
        /// </summary>
        private readonly bool _eraseFreeSectors;

        /// <summary>
        ///     Sectors ID entries in a FAT Sector
        /// </summary>
        private readonly int _fatSectorEntriesCount = 128;

        private readonly Queue<Sector> _flushingQueue = new Queue<Sector>(FlushingQueueSize);

        /// <summary>
        ///     Flag for sector recycling.
        /// </summary>
        private readonly bool _sectorRecycle;

        /// <summary>
        ///     True when update enabled
        /// </summary>
        private readonly UpdateMode _updateMode = UpdateMode.ReadOnly;

        internal int LockSectorId = -1;

        /// <summary>
        ///     Compound underlying stream. Null when new CF has been created.
        /// </summary>
        internal Stream SourceStream = null;

        internal bool TransactionLockAdded = false;

        internal bool TransactionLockAllocated = false;
        private byte[] _buffer = new byte[FlushingBufferMaxSize];

        /// <summary>
        ///     Contains a list with all the directory entries
        /// </summary>
        private List<IDirectoryEntry> _directoryEntries = new List<IDirectoryEntry>();

        /// <summary>
        ///     CompoundFile header
        /// </summary>
        private Header _header;

        /// <summary>
        ///     Used for thread safe locking
        /// </summary>
        private object _lockObject = new Object();

        /// <summary>
        ///     File sectors
        /// </summary>
        private SectorCollection _sectors = new SectorCollection();
        #endregion

        #region Properties
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
        public CFStorage RootStorage { get; private set; }

        /// <summary>
        ///     The root entry of all the <see cref="Directory" /> entries
        /// </summary>
        internal IDirectoryEntry RootEntry
        {
            get { return _directoryEntries[0]; }
        }

        /// <summary>
        ///     True when the compound file is closed
        /// </summary>
        internal bool IsClosed { get; private set; }

        /// <summary>
        ///     The name of the compound file, null when the compound file is opened from a stream
        /// </summary>
        public string FileName { get; private set; }
        #endregion

        #region Constructors
        /// <summary>
        ///     Create a blank, version 3 compound file.
        ///     Sector recycle is turned off to achieve the best reading/writing
        ///     performance in most common scenarios.
        /// </summary>
        /// <example>
        ///     <code>
        ///  
        ///      byte[] b = new byte[10000];
        ///      for (int i = 0; i &lt; 10000; i++)
        ///      {
        ///          b[i % 120] = (byte)i;
        ///      }
        /// 
        ///      CompoundFile cf = new CompoundFile();
        ///      CFStream myStream = cf.RootStorage.AddStream("MyStream");
        /// 
        ///      Assert.IsNotNull(myStream);
        ///      myStream.SetData(b);
        ///      cf.Save("MyCompoundFile.cfs");
        ///      cf.Close();
        ///      
        ///  </code>
        /// </example>
        public CompoundFile()
        {
            _header = new Header();
            _sectorRecycle = false;

            _difatSectorFATEntriesCount = (GetSectorSize()/4) - 1;
            _fatSectorEntriesCount = (GetSectorSize()/4);

            //Root -- 
            RootStorage = new CFStorage(this);

            RootStorage.DirEntry.SetEntryName("Root Entry");
            RootStorage.DirEntry.StgType = StgType.StgRoot;
        }

        /// <summary>
        ///     Create a new, blank, compound file.
        /// </summary>
        /// <param name="cfsVersion">Use a specific Compound File Version to set 512 or 4096 bytes sectors</param>
        /// <param name="sectorRecycle">If true, recycle unused sectors</param>
        /// <example>
        ///     <code>
        ///  
        ///      byte[] b = new byte[10000];
        ///      for (int i = 0; i &lt; 10000; i++)
        ///      {
        ///          b[i % 120] = (byte)i;
        ///      }
        /// 
        ///      CompoundFile cf = new CompoundFile(CFSVersion.Ver_4, true, true);
        ///      CFStream myStream = cf.RootStorage.AddStream("MyStream");
        /// 
        ///      Assert.IsNotNull(myStream);
        ///      myStream.SetData(b);
        ///      cf.Save("MyCompoundFile.cfs");
        ///      cf.Close();
        ///      
        ///  </code>
        /// </example>
        /// <remarks>
        ///     Sector recycling reduces data writing performances but avoids space wasting in scenarios with frequently
        ///     data manipulation of the same streams. The new compound file is open in Update mode.
        /// </remarks>
        public CompoundFile(CFSVersion cfsVersion, bool sectorRecycle)
        {
            _header = new Header((ushort) cfsVersion);
            _sectorRecycle = sectorRecycle;


            _difatSectorFATEntriesCount = (GetSectorSize()/4) - 1;
            _fatSectorEntriesCount = (GetSectorSize()/4);

            //Root -- 
            RootStorage = new CFStorage(this);

            RootStorage.DirEntry.SetEntryName("Root Entry");
            RootStorage.DirEntry.StgType = StgType.StgRoot;
        }

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
        ///     File will be open in read-only mode: it has to be saved
        ///     with a different filename. A wrapping implementation has to be provided
        ///     in order to remove/substitute an existing file. Version will be
        ///     automatically recognized from the file. Sector recycle is turned off
        ///     to achieve the best reading/writing performance in most common scenarios.
        /// </remarks>
        public CompoundFile(string fileName)
        {
            _sectorRecycle = false;
            _updateMode = UpdateMode.ReadOnly;
            _eraseFreeSectors = false;

            LoadFile(fileName);
            FileName = fileName;

            _difatSectorFATEntriesCount = (GetSectorSize()/4) - 1;
            _fatSectorEntriesCount = (GetSectorSize()/4);
        }

        /// <summary>
        ///     Load an existing compound file.
        /// </summary>
        /// <param name="fileName">Compound file to read from</param>
        /// <param name="sectorRecycle">If true, recycle unused sectors</param>
        /// <param name="updateMode">Select the update mode of the underlying data file</param>
        /// <param name="eraseFreeSectors">If true, overwrite with zeros unallocated sectors</param>
        /// <example>
        ///     <code>
        ///  string srcFilename = "data_YOU_CAN_CHANGE.xls";
        ///  
        ///  CompoundFile cf = new CompoundFile(srcFilename, UpdateMode.Update, true, true);
        /// 
        ///  Random r = new Random();
        /// 
        ///  byte[] buffer = GetBuffer(r.Next(3, 4095), 0x0A);
        /// 
        ///  cf.RootStorage.AddStream("MyStream").SetData(buffer);
        ///  
        ///  //This will persist data to the underlying media.
        ///  cf.Commit();
        ///  cf.Close();
        /// 
        ///  </code>
        /// </example>
        public CompoundFile(string fileName, UpdateMode updateMode, bool sectorRecycle, bool eraseFreeSectors)
        {
            _sectorRecycle = sectorRecycle;
            _updateMode = updateMode;
            _eraseFreeSectors = eraseFreeSectors;

            LoadFile(fileName);
            FileName = fileName;

            _difatSectorFATEntriesCount = (GetSectorSize()/4) - 1;
            _fatSectorEntriesCount = (GetSectorSize()/4);
        }

        /// <summary>
        ///     Load an existing compound file.
        /// </summary>
        /// <param name="stream">A stream containing a compound file to read</param>
        /// <param name="sectorRecycle">If true, recycle unused sectors</param>
        /// <param name="updateMode">Select the update mode of the underlying data file</param>
        /// <param name="eraseFreeSectors">If true, overwrite with zeros unallocated sectors</param>
        /// <example>
        ///     <code>
        ///  
        ///  string filename = "reportREAD.xls";
        ///    
        ///  FileStream fs = new FileStream(filename, FileMode.Open);
        ///  CompoundFile cf = new CompoundFile(fs, UpdateMode.ReadOnly, false, false);
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
        /// <exception cref="CFException">Raised when trying to open a non-seekable stream</exception>
        /// <exception cref="CFException">Raised stream is null</exception>
        public CompoundFile(Stream stream, UpdateMode updateMode, bool sectorRecycle, bool eraseFreeSectors)
        {
            _sectorRecycle = sectorRecycle;
            _updateMode = updateMode;
            _eraseFreeSectors = eraseFreeSectors;

            LoadStream(stream);

            _difatSectorFATEntriesCount = (GetSectorSize()/4) - 1;
            _fatSectorEntriesCount = (GetSectorSize()/4);
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
        /// <exception cref="CFException">Raised when trying to open a non-seekable stream</exception>
        /// <exception cref="CFException">Raised stream is null</exception>
        public CompoundFile(Stream stream)
        {
            LoadStream(stream);

            _difatSectorFATEntriesCount = (GetSectorSize()/4) - 1;
            _fatSectorEntriesCount = (GetSectorSize()/4);
        }
        #endregion

        #region Commit
        /// <summary>
        ///     Commit data changes since the previously commit operation
        ///     to the underlying supporting stream or file on the disk.
        /// </summary>
        /// <remarks>
        ///     This method can be used
        ///     only if the supporting stream has been opened in
        ///     <see cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.UpdateMode">Update mode</see>.
        /// </remarks>
        public void Commit()
        {
            Commit(false);
        }

        /// <summary>
        ///     Commit data changes since the previously commit operation
        ///     to the underlying supporting stream or file on the disk.
        /// </summary>
        /// <param name="releaseMemory">
        ///     If true, release loaded sectors to limit memory usage but reduces following read operations
        ///     performance
        /// </param>
        /// <remarks>
        ///     This method can be used only if
        ///     the supporting stream has been opened in
        ///     <see cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.UpdateMode">Update mode</see>.
        /// </remarks>
        public void Commit(bool releaseMemory)
        {
            if (IsClosed)
                throw new CFDisposedException("Compound File closed: cannot commit data");

            if (_updateMode != UpdateMode.Update)
                throw new CFInvalidOperation("Cannot commit data in Read-Only update mode");

            var sId = -1;
            int sectorCount;
            int bufOffset;
            var sSize = GetSectorSize();

            if (_header.MajorVersion != (ushort) CFSVersion.Ver3)
                CheckForLockSector();

            SourceStream.Seek(0, SeekOrigin.Begin);
            SourceStream.Write(new byte[GetSectorSize()], 0, sSize);

            CommitDirectory();

            var gap = true;

            for (var i = 0; i < _sectors.Count; i++)
            {
                var sector = _sectors[i];

                if (sector != null && sector.DirtyFlag && _flushingQueue.Count < _buffer.Length/sSize)
                {
                    //First of a block of contiguous sectors, mark id, start enqueuing
                    if (gap)
                    {
                        sId = sector.Id;
                        gap = false;
                    }

                    _flushingQueue.Enqueue(sector);
                }
                else
                {
                    //Found a gap, stop enqueuing, flush a write operation

                    gap = true;
                    sectorCount = _flushingQueue.Count;

                    if (sectorCount == 0) continue;

                    bufOffset = 0;
                    while (_flushingQueue.Count > 0)
                    {
                        var r = _flushingQueue.Dequeue();
                        Buffer.BlockCopy(r.GetData(), 0, _buffer, bufOffset, sSize);
                        r.DirtyFlag = false;

                        if (releaseMemory)
                            r.ReleaseData();

                        bufOffset += sSize;
                    }

                    SourceStream.Seek((sSize + sId*(long) sSize), SeekOrigin.Begin);
                    SourceStream.Write(_buffer, 0, sectorCount*sSize);
                }
            }

            sectorCount = _flushingQueue.Count;
            bufOffset = 0;

            while (_flushingQueue.Count > 0)
            {
                var sector = _flushingQueue.Dequeue();
                Buffer.BlockCopy(sector.GetData(), 0, _buffer, bufOffset, sSize);
                sector.DirtyFlag = false;

                if (releaseMemory)
                    sector.ReleaseData();

                bufOffset += sSize;
            }

            if (sectorCount != 0)
            {
                SourceStream.Seek(sSize + sId*(long) sSize, SeekOrigin.Begin);
                SourceStream.Write(_buffer, 0, sectorCount*sSize);
            }

            // Seek to beginning position and save header (first 512 or 4096 bytes)
            SourceStream.Seek(0, SeekOrigin.Begin);
            _header.Write(SourceStream);
        }
        #endregion

        #region Load
        /// <summary>
        ///     Load compound file from an existing stream.
        /// </summary>
        /// <param name="stream">Stream to load compound file from</param>
        private void Load(Stream stream)
        {
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

                RootStorage = new CFStorage(this, _directoryEntries[0]);
            }
            catch (Exception)
            {
                if (stream != null)
                    stream.Close();

                throw;
            }
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

        #region LoadFile
        /// <summary>
        ///     Loads a compound file from a file
        /// </summary>
        /// <param name="fileName"></param>
        private void LoadFile(string fileName)
        {
            FileStream fileStream = null;

            try
            {
                fileStream = _updateMode == UpdateMode.ReadOnly
                    ? new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read)
                    : new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);

                Load(fileStream);
            }
            catch (Exception)
            {
                if (fileStream != null)
                    fileStream.Close();


                throw;
            }
        }
        #endregion

        #region LoadStream
        /// <summary>
        ///     Loads a compound file from a <see cref="stream" />
        /// </summary>
        /// <param name="stream"></param>
        /// <exception cref="CFException">Raised when the stream is null or non-seekable</exception>
        private void LoadStream(Stream stream)
        {
            if (stream == null)
                throw new CFException("Stream parameter cannot be null");

            if (!stream.CanSeek)
                throw new CFException("Cannot load a non-seekable Stream");


            stream.Seek(0, SeekOrigin.Begin);

            Load(stream);
        }
        #endregion

        #region Save
        /// <summary>
        ///     Saves the in-memory image of Compound File to a file.
        /// </summary>
        /// <param name="fileName">File name to write the compound file to</param>
        /// <exception cref="CFException">Raised when file is closed</exception>
        public void Save(string fileName)
        {
            if (IsClosed)
                throw new CFException("Compound file closed: cannot save data");

            FileStream fileStream = null;

            try
            {
                fileStream = new FileStream(fileName, FileMode.Create);
                Save(fileStream);
            }
            catch (Exception ex)
            {
                throw new CFException("Error saving file [" + fileName + "]", ex);
            }
            finally
            {
                if (fileStream != null)
                    fileStream.Flush();

                if (fileStream != null)
                    fileStream.Close();
            }
        }

        /// <summary>
        ///     Saves the in-memory image of Compound File to a stream.
        /// </summary>
        /// <remarks>
        ///     Destination Stream must be seekable.
        /// </remarks>
        /// <param name="stream">The stream to save compound File to</param>
        /// <exception cref="CFException">Raised if destination stream is not seekable</exception>
        /// <exception cref="CFDisposedException">Raised if Compound File Storage has been already disposed</exception>
        /// <example>
        ///     <code>
        ///     MemoryStream ms = new MemoryStream(size);
        /// 
        ///     CompoundFile cf = new CompoundFile();
        ///     CFStorage st = cf.RootStorage.AddStorage("MyStorage");
        ///     CFStream sm = st.AddStream("MyStream");
        /// 
        ///     byte[] b = new byte[]{0x00,0x01,0x02,0x03};
        /// 
        ///     sm.SetData(b);
        ///     cf.Save(ms);
        ///     cf.Close();
        ///  </code>
        /// </example>
        public void Save(Stream stream)
        {
            if (IsClosed)
                throw new CFDisposedException("Compound File closed: cannot save data");

            if (!stream.CanSeek)
                throw new CFException("Cannot save on a non-seekable stream");

            CheckForLockSector();
            var sSize = GetSectorSize();

            try
            {
                stream.Write(new byte[sSize], 0, sSize);

                CommitDirectory();

                for (var i = 0; i < _sectors.Count; i++)
                {
                    var sector = _sectors[i] ?? new Sector(sSize, SourceStream) {Id = i};
                    stream.Write(sector.GetData(), 0, sSize);
                }

                stream.Seek(0, SeekOrigin.Begin);
                _header.Write(stream);
            }
            catch (Exception ex)
            {
                throw new CFException("Internal error while saving compound file to stream ", ex);
            }
        }
        #endregion

        #region Close
        /// <summary>
        ///     Close the Compound File object <see cref="CompoundFile">CompoundFile</see> and
        ///     free all associated resources (e.g. open file handle and allocated memory).
        ///     <remarks>
        ///         When the <see cref="CompoundFile.Close()">Close</see> method is called,
        ///         all the associated stream and storage objects are invalidated:
        ///         any operation invoked on them will produce a
        ///         <see cref="CFDisposedException">CFDisposedException</see>.
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

        #region HasSourceStream
        /// <summary>
        ///     Return true if this compound file has been loaded from an existing file or stream
        /// </summary>
        public bool HasSourceStream
        {
            get { return SourceStream != null; }
        }
        #endregion

        #region SetMiniSectorChain
        /// <summary>
        ///     Allocate space, setup sectors id and refresh header
        ///     for the new or updated mini sector chain.
        /// </summary>
        /// <param name="sectorChain">The new MINI sector chain</param>
        private void SetMiniSectorChain(IList<Sector> sectorChain)
        {
            var miniFAT
                = GetSectorChain(_header.FirstMiniFATSectorId, SectorType.Normal);

            var miniStream
                = GetSectorChain(RootEntry.StartSector, SectorType.Normal);

            var miniFATView
                = new StreamView(
                    miniFAT,
                    GetSectorSize(),
                    _header.MiniFATSectorsNumber*Sector.MinisectorSize,
                    SourceStream
                    );

            var miniStreamView
                = new StreamView(
                    miniStream,
                    GetSectorSize(),
                    RootStorage.Size,
                    SourceStream);

            // Set updated/new sectors within the ministream
            foreach (var sector in sectorChain)
            {
                if (sector.Id != -1)
                {
                    // Overwrite
                    miniStreamView.Seek(Sector.MinisectorSize*sector.Id, SeekOrigin.Begin);
                    miniStreamView.Write(sector.GetData(), 0, Sector.MinisectorSize);
                }
                else
                {
                    // Allocate, position ministream at the end of already allocated
                    // ministream's sectors

                    miniStreamView.Seek(RootStorage.Size, SeekOrigin.Begin);
                    miniStreamView.Write(sector.GetData(), 0, Sector.MinisectorSize);
                    sector.Id = (int) (miniStreamView.Position - Sector.MinisectorSize)/Sector.MinisectorSize;

                    RootStorage.DirEntry.Size = miniStreamView.Length;
                }
            }

            // Update miniFAT
            for (var i = 0; i < sectorChain.Count - 1; i++)
            {
                var currentId = sectorChain[i].Id;
                var nextId = sectorChain[i + 1].Id;

                //AssureLength(miniFATView, Math.Max(currentId * SIZE_OF_SID, nextId * SIZE_OF_SID));

                miniFATView.Seek(currentId*4, SeekOrigin.Begin);
                miniFATView.Write(BitConverter.GetBytes(nextId), 0, 4);
            }

            // Write End of Chain in MiniFAT
            miniFATView.Seek(sectorChain[sectorChain.Count - 1].Id*SizeOfSID, SeekOrigin.Begin);
            miniFATView.Write(BitConverter.GetBytes(Sector.Endofchain), 0, 4);

            // Update sector chains
            SetNormalSectorChain(miniStreamView.BaseSectorChain);
            SetNormalSectorChain(miniFATView.BaseSectorChain);

            //Update HEADER and root storage when ministream changes
            if (miniFAT.Count > 0)
            {
                RootStorage.DirEntry.StartSector = miniStream[0].Id;
                _header.MiniFATSectorsNumber = (uint) miniFAT.Count;
                _header.FirstMiniFATSectorId = miniFAT[0].Id;
            }
        }
        #endregion

        #region FreeChain
        private void FreeChain(IList<Sector> sectorChain, bool zeroSector)
        {
            var fat = GetSectorChain(-1, SectorType.FAT);

            var fatView = new StreamView(fat, GetSectorSize(), fat.Count*GetSectorSize(), SourceStream);

            // Zeroes out sector data (if requested)
            if (zeroSector)
            {
                foreach (var sector in sectorChain)
                    sector.ZeroData();
            }

            // Update FAT marking unallocated sectors
            for (var i = 0; i < sectorChain.Count - 1; i++)
            {
                var currentId = sectorChain[i].Id;

                //AssureLength(FATView, Math.Max(currentId * SIZE_OF_SID, nextId * SIZE_OF_SID));

                fatView.Seek(currentId*4, SeekOrigin.Begin);
                fatView.Write(BitConverter.GetBytes(Sector.FreeSector), 0, 4);
            }
        }
        #endregion

        #region FreeMiniChain
        private void FreeMiniChain(IList<Sector> sectorChain, bool zeroSector)
        {
            var zeroedMiniSector = new byte[Sector.MinisectorSize];

            var miniFAT
                = GetSectorChain(_header.FirstMiniFATSectorId, SectorType.Normal);

            var miniStream = GetSectorChain(RootEntry.StartSector, SectorType.Normal);

            var miniFATView = new StreamView(miniFAT, GetSectorSize(),
                _header.MiniFATSectorsNumber*Sector.MinisectorSize,
                SourceStream);

            var miniStreamView = new StreamView(miniStream, GetSectorSize(), RootStorage.Size, SourceStream);

            // Set updated/new sectors within the ministream
            if (zeroSector)
            {
                foreach (var sector in sectorChain)
                {
                    if (sector.Id == -1) continue;
                    // Overwrite
                    miniStreamView.Seek(Sector.MinisectorSize*sector.Id, SeekOrigin.Begin);
                    miniStreamView.Write(zeroedMiniSector, 0, Sector.MinisectorSize);
                }
            }

            // Update miniFAT
            for (var i = 0; i < sectorChain.Count - 1; i++)
            {
                var currentId = sectorChain[i].Id;
                miniFATView.Seek(currentId*4, SeekOrigin.Begin);
                miniFATView.Write(BitConverter.GetBytes(Sector.FreeSector), 0, 4);
            }

            //AssureLength(miniFATView, sectorChain[sectorChain.Count - 1].Id * SIZE_OF_SID);

            // Write End of Chain in MiniFAT
            miniFATView.Seek(sectorChain[sectorChain.Count - 1].Id*SizeOfSID, SeekOrigin.Begin);
            miniFATView.Write(BitConverter.GetBytes(Sector.FreeSector), 0, 4);

            // Update sector chains
            SetNormalSectorChain(miniStreamView.BaseSectorChain);
            SetNormalSectorChain(miniFATView.BaseSectorChain);

            //Update HEADER and root storage when ministream changes
            if (miniFAT.Count > 0)
            {
                RootStorage.DirEntry.StartSector = miniStream[0].Id;
                _header.MiniFATSectorsNumber = (uint) miniFAT.Count;
                _header.FirstMiniFATSectorId = miniFAT[0].Id;
            }
        }
        #endregion

        #region SetNormalSectorChain
        /// <summary>
        ///     Allocate space, setup sectors id and refresh header
        ///     for the new or updated sector chain.
        /// </summary>
        /// <param name="sectorChain">The new or updated generic sector chain</param>
        private void SetNormalSectorChain(List<Sector> sectorChain)
        {
            foreach (var s in sectorChain)
            {
                if (s.Id != -1) continue;
                _sectors.Add(s);
                s.Id = _sectors.Count - 1;
            }

            SetFATSectorChain(sectorChain);
        }
        #endregion

        #region CheckForLockSector
        /// <summary>
        ///     Check for transaction lock sector addition and mark it in the FAT.
        /// </summary>
        private void CheckForLockSector()
        {
            //If transaction lock has been added and not yet allocated in the FAT...
            if (TransactionLockAdded && !TransactionLockAllocated)
            {
                var fatStream = new StreamView(GetFatSectorChain(), GetSectorSize(), SourceStream);

                fatStream.Seek(LockSectorId*4, SeekOrigin.Begin);
                fatStream.Write(BitConverter.GetBytes(Sector.Endofchain), 0, 4);

                TransactionLockAllocated = true;
            }
        }
        #endregion

        #region SetFATSectorChain
        /// <summary>
        ///     Allocate space, setup sectors id and refresh header
        ///     for the new or updated FAT sector chain.
        /// </summary>
        /// <param name="sectorChain">The new or updated generic sector chain</param>
        private void SetFATSectorChain(IList<Sector> sectorChain)
        {
            var fatSectors = GetSectorChain(-1, SectorType.FAT);
            var fatStream =
                new StreamView(
                    fatSectors,
                    GetSectorSize(),
                    _header.FATSectorsNumber*GetSectorSize(), SourceStream
                    );

            // Write FAT chain values --

            for (var i = 0; i < sectorChain.Count - 1; i++)
            {
                var sN = sectorChain[i + 1];
                var sC = sectorChain[i];

                fatStream.Seek(sC.Id*4, SeekOrigin.Begin);
                fatStream.Write(BitConverter.GetBytes(sN.Id), 0, 4);
            }

            fatStream.Seek(sectorChain[sectorChain.Count - 1].Id*4, SeekOrigin.Begin);
            fatStream.Write(BitConverter.GetBytes(Sector.Endofchain), 0, 4);

            // Merge chain to CFS
            SetDIFATSectorChain(fatStream.BaseSectorChain);
        }
        #endregion

        #region SetDIFATSectorChain
        /// <summary>
        ///     Setup the DIFAT sector chain
        /// </summary>
        /// <param name="faTsectorChain">A FAT sector chain</param>
        private void SetDIFATSectorChain(List<Sector> faTsectorChain)
        {
            // Get initial sector's count
            _header.FATSectorsNumber = faTsectorChain.Count;

            // Allocate Sectors
            foreach (var s in faTsectorChain)
            {
                if (s.Id != -1) continue;
                _sectors.Add(s);
                s.Id = _sectors.Count - 1;
                s.Type = SectorType.FAT;
            }

            // Sector count...
            var nCurrentSectors = _sectors.Count;

            // Temp DIFAT count
            var nDIFATSectors = (int) _header.DIFATSectorsNumber;

            if (faTsectorChain.Count > HeaderDIFATEntriesCount)
            {
                nDIFATSectors =
                    Ceiling((double) (faTsectorChain.Count - HeaderDIFATEntriesCount)/_difatSectorFATEntriesCount);
                nDIFATSectors = LowSaturation(nDIFATSectors - (int) _header.DIFATSectorsNumber); //required DIFAT
            }

            // ...sum with new required DIFAT sectors count
            nCurrentSectors += nDIFATSectors;

            // ReCheck FAT bias
            while (_header.FATSectorsNumber*_fatSectorEntriesCount < nCurrentSectors)
            {
                var extraFATSector = new Sector(GetSectorSize(), SourceStream);
                _sectors.Add(extraFATSector);

                extraFATSector.Id = _sectors.Count - 1;
                extraFATSector.Type = SectorType.FAT;

                faTsectorChain.Add(extraFATSector);

                _header.FATSectorsNumber++;
                nCurrentSectors++;

                //... so, adding a FAT sector may induce DIFAT sectors to increase by one
                // and consequently this may induce ANOTHER FAT sector (TO-THINK: May this condition occure ?)
                if (nDIFATSectors*_difatSectorFATEntriesCount >= (_header.FATSectorsNumber > HeaderDIFATEntriesCount
                    ? _header.FATSectorsNumber - HeaderDIFATEntriesCount
                    : 0)) continue;
                nDIFATSectors++;
                nCurrentSectors++;
            }

            var difatSectors = GetSectorChain(-1, SectorType.DIFAT);

            var difatStream = new StreamView(difatSectors, GetSectorSize(), SourceStream);

            // Write DIFAT Sectors (if required)
            // Save room for the following chaining
            for (var i = 0; i < faTsectorChain.Count; i++)
            {
                if (i < HeaderDIFATEntriesCount)
                    _header.DIFAT[i] = faTsectorChain[i].Id;
                else
                {
                    // room for DIFAT chaining at the end of any DIFAT sector (4 bytes)
                    if (i != HeaderDIFATEntriesCount &&
                        (i - HeaderDIFATEntriesCount)%_difatSectorFATEntriesCount == 0)
                    {
                        var temp = new byte[sizeof (int)];
                        difatStream.Write(temp, 0, sizeof (int));
                    }

                    difatStream.Write(BitConverter.GetBytes(faTsectorChain[i].Id), 0, sizeof (int));
                }
            }

            // Allocate room for DIFAT sectors
            foreach (var sector in difatStream.BaseSectorChain)
            {
                if (sector.Id != -1) continue;
                _sectors.Add(sector);
                sector.Id = _sectors.Count - 1;
                sector.Type = SectorType.DIFAT;
            }

            _header.DIFATSectorsNumber = (uint) nDIFATSectors;


            // Chain first sector
            if (difatStream.BaseSectorChain != null && difatStream.BaseSectorChain.Count > 0)
            {
                _header.FirstDIFATSectorId = difatStream.BaseSectorChain[0].Id;

                // Update header information
                _header.DIFATSectorsNumber = (uint) difatStream.BaseSectorChain.Count;

                // Write chaining information at the end of DIFAT Sectors
                for (var i = 0; i < difatStream.BaseSectorChain.Count - 1; i++)
                {
                    Buffer.BlockCopy(
                        BitConverter.GetBytes(difatStream.BaseSectorChain[i + 1].Id),
                        0,
                        difatStream.BaseSectorChain[i].GetData(),
                        GetSectorSize() - sizeof (int),
                        4);
                }

                Buffer.BlockCopy(
                    BitConverter.GetBytes(Sector.Endofchain),
                    0,
                    difatStream.BaseSectorChain[difatStream.BaseSectorChain.Count - 1].GetData(),
                    GetSectorSize() - sizeof (int),
                    sizeof (int)
                    );
            }
            else
                _header.FirstDIFATSectorId = Sector.Endofchain;

            // Mark DIFAT Sectors in FAT
            var fatSv = new StreamView(faTsectorChain, GetSectorSize(), _header.FATSectorsNumber*GetSectorSize(),
                SourceStream);

            for (var i = 0; i < _header.DIFATSectorsNumber; i++)
            {
                fatSv.Seek(difatStream.BaseSectorChain[i].Id*4, SeekOrigin.Begin);
                fatSv.Write(BitConverter.GetBytes(Sector.DifSector), 0, 4);
            }

            for (var i = 0; i < _header.FATSectorsNumber; i++)
            {
                fatSv.Seek(fatSv.BaseSectorChain[i].Id*4, SeekOrigin.Begin);
                fatSv.Write(BitConverter.GetBytes(Sector.FATSector), 0, 4);
            }

            _header.FATSectorsNumber = fatSv.BaseSectorChain.Count;
        }
        #endregion

        #region GetDifatSectorChain
        /// <summary>
        ///     Get the DIFAT Sector chain
        /// </summary>
        /// <returns>A list of DIFAT sectors</returns>
        /// <exception cref="CFCorruptedFileException">Raised when DIFAT sectors count is mismatched</exception>
        private List<Sector> GetDifatSectorChain()
        {
            var result
                = new List<Sector>();

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
                var nextSecId = BitConverter.ToInt32(sector.GetData(), GetSectorSize() - 4);

                // Strictly speaking, the following condition is not correct from
                // a specification point of view:
                // only ENDOFCHAIN should break DIFAT chain but 
                // a lot of existing compound files use FREESECT as DIFAT chain termination
                if (nextSecId == Sector.FreeSector || nextSecId == Sector.Endofchain) break;

                validationCount--;

                if (validationCount < 0)
                {
                    Close();
                    throw new CFCorruptedFileException("DIFAT sectors count mismatched. Corrupted compound file");
                }

                sector = _sectors[nextSecId];

                if (sector == null)
                {
                    sector = new Sector(GetSectorSize(), SourceStream) {Id = nextSecId};
                    _sectors[nextSecId] = sector;
                }

                result.Add(sector);
            }

            return result;
        }
        #endregion

        #region GetFatSectorChain
        /// <summary>
        ///     Get the FAT sector chain
        /// </summary>
        /// <returns>List of FAT sectors</returns>
        private List<Sector> GetFatSectorChain()
        {
            const int numberOfHeaderFATEntry = 109; //Number of FAT sectors id in the header

            var result
                = new List<Sector>();

            int nextSecId;

            var difatSectors = GetDifatSectorChain();

            var idx = 0;

            // Read FAT entries from the header Fat entry array (max 109 entries)
            while (idx < _header.FATSectorsNumber && idx < numberOfHeaderFATEntry)
            {
                nextSecId = _header.DIFAT[idx];
                var sector = _sectors[nextSecId];

                if (sector == null)
                {
                    sector = new Sector(GetSectorSize(), SourceStream) {Id = nextSecId, Type = SectorType.FAT};
                    _sectors[nextSecId] = sector;
                }

                result.Add(sector);

                idx++;
            }

            //Is there any DIFAT sector containing other FAT entries ?
            if (difatSectors.Count <= 0) return result;
            var difatStream
                = new StreamView
                    (
                    difatSectors,
                    GetSectorSize(),
                    _header.FATSectorsNumber > numberOfHeaderFATEntry
                        ? (_header.FATSectorsNumber - numberOfHeaderFATEntry)*4
                        : 0,
                    SourceStream
                    );

            var nextDIFATSectorBuffer = new byte[4];

            difatStream.Read(nextDIFATSectorBuffer, 0, 4);
            nextSecId = BitConverter.ToInt32(nextDIFATSectorBuffer, 0);

            var i = 0;
            var numberOfFatHeaderEntries = numberOfHeaderFATEntry;

            while (numberOfFatHeaderEntries < _header.FATSectorsNumber)
            {
                if (difatStream.Position == ((GetSectorSize() - 4) + i*GetSectorSize()))
                {
                    difatStream.Seek(4, SeekOrigin.Current);
                    i++;
                    continue;
                }

                var sector = _sectors[nextSecId];

                if (sector == null)
                {
                    sector = new Sector(GetSectorSize(), SourceStream) {Type = SectorType.FAT, Id = nextSecId};
                    _sectors[nextSecId] = sector; //UUU
                }

                result.Add(sector);

                difatStream.Read(nextDIFATSectorBuffer, 0, 4);
                nextSecId = BitConverter.ToInt32(nextDIFATSectorBuffer, 0);
                numberOfFatHeaderEntries++;
            }

            return result;
        }
        #endregion

        #region GetNormalSectorChain
        /// <summary>
        ///     Get a standard sector chain
        /// </summary>
        /// <param name="secId">First SecID of the required chain</param>
        /// <returns>A list of sectors</returns>
        /// <exception cref="CFCorruptedFileException">Raised when the file is corrupt</exception>
        private List<Sector> GetNormalSectorChain(int secId)
        {
            var result = new List<Sector>();

            var nextSecId = secId;

            var fatSectors = GetFatSectorChain();

            var fatStream
                = new StreamView(fatSectors, GetSectorSize(), fatSectors.Count*GetSectorSize(), SourceStream);

            while (true)
            {
                if (nextSecId == Sector.Endofchain) break;

                if (nextSecId >= _sectors.Count)
                    throw new CFCorruptedFileException(
                        string.Format(
                            "Next Sector ID reference an out of range sector. NextID : {0} while sector count {1}",
                            nextSecId, _sectors.Count));

                var sector = _sectors[nextSecId];
                if (sector == null)
                {
                    sector = new Sector(GetSectorSize(), SourceStream) {Id = nextSecId, Type = SectorType.Normal};
                    _sectors[nextSecId] = sector;
                }

                result.Add(sector);

                fatStream.Seek(nextSecId*4, SeekOrigin.Begin);
                var next = fatStream.ReadInt32();

                if (next != nextSecId)
                    nextSecId = next;
                else
                    throw new CFCorruptedFileException("Cyclic sector chain found. File is corrupted");
            }

            return result;
        }
        #endregion

        #region GetMiniSectorChain
        /// <summary>
        ///     Get a mini sector chain
        /// </summary>
        /// <param name="sectorId">First sector id of the required chain</param>
        /// <returns>A list of mini sectors (64 bytes)</returns>
        private List<Sector> GetMiniSectorChain(int sectorId)
        {
            var result = new List<Sector>();

            if (sectorId == Sector.Endofchain) return result;
            var miniFAT = GetNormalSectorChain(_header.FirstMiniFATSectorId);
            var miniStream = GetNormalSectorChain(RootEntry.StartSector);

            var miniFATView
                = new StreamView(miniFAT, GetSectorSize(), _header.MiniFATSectorsNumber*Sector.MinisectorSize,
                    SourceStream);

            var miniStreamView =
                new StreamView(miniStream, GetSectorSize(), RootStorage.Size, SourceStream);

            var miniFATReader = new BinaryReader(miniFATView);

            var nextSectorId = sectorId;

            while (true)
            {
                if (nextSectorId == Sector.Endofchain)
                    break;

                var miniSector = new Sector(Sector.MinisectorSize, SourceStream)
                {
                    Id = nextSectorId,
                    Type = SectorType.Mini
                };

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
        ///     Get a sector chain from a compound file given the first sector ID
        ///     and the required sector type.
        /// </summary>
        /// <param name="sectorId">First chain sector's id </param>
        /// <param name="chainType">Type of Sectors in the required chain (mini sectors, normal sectors or FAT)</param>
        /// <returns>A list of Sectors as the result of their concatenation</returns>
        internal List<Sector> GetSectorChain(int sectorId, SectorType chainType)
        {
            switch (chainType)
            {
                case SectorType.DIFAT:
                    return GetDifatSectorChain();

                case SectorType.FAT:
                    return GetFatSectorChain();

                case SectorType.Normal:
                    return GetNormalSectorChain(sectorId);

                case SectorType.Mini:
                    return GetMiniSectorChain(sectorId);

                default:
                    throw new CFException("Unsupproted chain type");
            }
        }
        #endregion

        #region CFSVersion
        /// <summary>
        ///     Returns the version number of the compound file storage
        /// </summary>
        public CFSVersion Version
        {
            get { return (CFSVersion) _header.MajorVersion; }
        }
        #endregion

        #region InsertNewDirectoryEntry
        /// <summary>
        ///     Inserts a new <see cref="directoryEntry" />
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
                    if (_directoryEntries[i].StgType == StgType.StgInvalid)
                    {
                        _directoryEntries[i] = directoryEntry;
                        directoryEntry.SID = i;
                        return;
                    }
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
        ///     Returns the children tree for the given <see cref="sid" />
        /// </summary>
        /// <param name="sid"></param>
        /// <returns></returns>
        internal BinarySearchTree<CFItem> GetChildrenTree(int sid)
        {
            var binarySearchTree = new BinarySearchTree<CFItem>(new CFItemComparer());

            // Load children from their original tree.
            DoLoadChildren(binarySearchTree, _directoryEntries[sid]);

            // Rebuild of (Red)-Black tree of entry children.
            binarySearchTree.VisitTreeInOrder(RefreshSIDs);

            return binarySearchTree;
        }
        #endregion

        #region LoadChildren
        private void DoLoadChildren(ICollection<CFItem> bst, IDirectoryEntry directoryEntry)
        {
            if (directoryEntry.Child == DirectoryEntry.Nostream) return;
            if (_directoryEntries[directoryEntry.Child].StgType == StgType.StgInvalid) return;

            if (_directoryEntries[directoryEntry.Child].StgType == StgType.StgStream)
                bst.Add(new CFStream(this, _directoryEntries[directoryEntry.Child]));
            else
                bst.Add(new CFStorage(this, _directoryEntries[directoryEntry.Child]));

            LoadSiblings(bst, _directoryEntries[directoryEntry.Child]);
        }

        /// <summary>
        ///     Doubling methods allows iterative behavior while avoiding to insert duplicate items
        /// </summary>
        /// <param name="binarySearchTree"></param>
        /// <param name="directoryEntry"></param>
        private void LoadSiblings(ICollection<CFItem> binarySearchTree, IDirectoryEntry directoryEntry)
        {
            if (directoryEntry.LeftSibling != DirectoryEntry.Nostream)
            {
                // If there're more left siblings load them...
                DoLoadSiblings(binarySearchTree, _directoryEntries[directoryEntry.LeftSibling]);
            }

            if (directoryEntry.RightSibling != DirectoryEntry.Nostream)
            {
                // If there're more right siblings load them...
                DoLoadSiblings(binarySearchTree, _directoryEntries[directoryEntry.RightSibling]);
            }
        }

        private void DoLoadSiblings(ICollection<CFItem> binarySearchTree, IDirectoryEntry directoryEntry)
        {
            while (true)
            {
                if (ValidateSibling(directoryEntry.LeftSibling))
                {
                    // If there're more left siblings load them...
                    DoLoadSiblings(binarySearchTree, _directoryEntries[directoryEntry.LeftSibling]);
                }

                switch (_directoryEntries[directoryEntry.SID].StgType)
                {
                    case StgType.StgStream:
                        binarySearchTree.Add(new CFStream(this, _directoryEntries[directoryEntry.SID]));
                        break;
                    case StgType.StgStorage:
                        binarySearchTree.Add(new CFStorage(this, _directoryEntries[directoryEntry.SID]));
                        break;
                }


                if (ValidateSibling(directoryEntry.RightSibling))
                {
                    // If there're more right siblings load them...
                    directoryEntry = _directoryEntries[directoryEntry.RightSibling];
                    continue;
                }
                break;
            }
        }
        #endregion

        #region ValidateSibling
        /// <summary>
        ///     Validates all the siblings
        /// </summary>
        /// <param name="sid"></param>
        /// <returns></returns>
        /// <exception cref="CFCorruptedFileException">Raised when there is an invalid reference of storage type</exception>
        private bool ValidateSibling(int sid)
        {
            if (sid == DirectoryEntry.Nostream) return false;
            // if this siblings id does not overflow current list
            if (sid >= _directoryEntries.Count)
                return false;

            //if this sibling is valid...
            if (_directoryEntries[sid].StgType == StgType.StgInvalid)
            {
                Close();
                throw new CFCorruptedFileException(
                    "A directory entry has a valid reference to an invalid storage type directory");
            }

            if (Enum.IsDefined(typeof (StgType), _directoryEntries[sid].StgType))
                return true; //No fault condition encountered for sid being validated
            Close();
            throw new CFCorruptedFileException("A directory entry has an invalid storage type");
        }
        #endregion

        #region LoadDirectories
        /// <summary>
        ///     Load directory entries from compound file. Header and FAT MUST be already loaded.
        /// </summary>
        private void LoadDirectories()
        {
            var directoryChain
                = GetSectorChain(_header.FirstDirectorySectorId, SectorType.Normal);

            if (_header.FirstDirectorySectorId == Sector.Endofchain)
                _header.FirstDirectorySectorId = directoryChain[0].Id;

            var dirReader
                = new StreamView(directoryChain, GetSectorSize(), directoryChain.Count*GetSectorSize(), SourceStream);


            while (dirReader.Position < directoryChain.Count*GetSectorSize())
            {
                var de
                    = new DirectoryEntry(StgType.StgInvalid);

                //We are not inserting dirs. Do not use 'InsertNewDirectoryEntry'
                de.Read(dirReader);
                _directoryEntries.Add(de);
                de.SID = _directoryEntries.Count - 1;
            }
        }
        #endregion

        #region RemoveDirectoryEntry
        /// <summary>
        ///     Removes an directory entry
        /// </summary>
        /// <param name="sid"></param>
        /// <exception cref="CFException">Raised when the <see cref="sid" /> is invalid</exception>
        internal void RemoveDirectoryEntry(int sid)
        {
            if (sid >= _directoryEntries.Count)
                throw new CFException("Invalid SID of the directory entry to remove");

            if (_directoryEntries[sid].StgType == StgType.StgStream)
            {
                // Clear the associated stream (or ministream) if required
                if (_directoryEntries[sid].Size > 0) //thanks to Mark Bosold for this !
                {
                    if (_directoryEntries[sid].Size < _header.MinSizeStandardStream)
                    {
                        var miniChain
                            = GetSectorChain(_directoryEntries[sid].StartSector, SectorType.Mini);
                        FreeMiniChain(miniChain, _eraseFreeSectors);
                    }
                    else
                    {
                        var chain
                            = GetSectorChain(_directoryEntries[sid].StartSector, SectorType.Normal);
                        FreeChain(chain, _eraseFreeSectors);
                    }
                }
            }


            var r = new Random();
            _directoryEntries[sid].SetEntryName("_DELETED_NAME_" + r.Next(short.MaxValue));
            _directoryEntries[sid].StgType = StgType.StgInvalid;
        }
        #endregion

        #region CommitDirectory
        /// <summary>
        ///     Commit directory entries change on the Current Source stream
        /// </summary>
        private void CommitDirectory()
        {
            const int directorySize = 128;

            var directorySectors
                = GetSectorChain(_header.FirstDirectorySectorId, SectorType.Normal);

            var sv = new StreamView(directorySectors, GetSectorSize(), 0, SourceStream);

            foreach (var directoryEntry in _directoryEntries)
                directoryEntry.Write(sv);

            var delta = _directoryEntries.Count;

            while (delta%(GetSectorSize()/directorySize) != 0)
            {
                var dummy = new DirectoryEntry(StgType.StgInvalid);
                dummy.Write(sv);
                delta++;
            }

            foreach (var s in directorySectors)
            {
                s.Type = SectorType.Directory;
            }

            SetNormalSectorChain(directorySectors);

            _header.FirstDirectorySectorId = directorySectors[0].Id;

            // Version 4 supports directory sectors count
            _header.DirectorySectorsNumber = _header.MajorVersion == 3 ? 0 : directorySectors.Count;
        }
        #endregion

        #region RefreshSIDs
        /// <summary>
        ///     Refreshes all SID's for the give node
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

        #region FindFreeSectors
        /// <summary>
        ///     Scan FAT o miniFAT for free sectors to reuse.
        /// </summary>
        /// <param name="sType">Type of sector to look for</param>
        /// <returns>A stack of available sectors or minisectors already allocated</returns>
        internal Queue<Sector> FindFreeSectors(SectorType sType)
        {
            var freeList = new Queue<Sector>();

            if (sType == SectorType.Normal)
            {
                var fatChain = GetSectorChain(-1, SectorType.FAT);
                var fatStream = new StreamView(fatChain, GetSectorSize(), _header.FATSectorsNumber*GetSectorSize(),
                    SourceStream);

                var idx = 0;

                while (idx < _sectors.Count)
                {
                    var id = fatStream.ReadInt32();

                    if (id == Sector.FreeSector)
                    {
                        if (_sectors[idx] == null)
                        {
                            var sector = new Sector(GetSectorSize(), SourceStream) {Id = idx};
                            _sectors[idx] = sector;
                        }

                        freeList.Enqueue(_sectors[idx]);
                    }

                    idx++;
                }
            }
            else
            {
                var miniFAT
                    = GetSectorChain(_header.FirstMiniFATSectorId, SectorType.Normal);

                var miniFATView
                    = new StreamView(miniFAT, GetSectorSize(), _header.MiniFATSectorsNumber*Sector.MinisectorSize,
                        SourceStream);

                var miniStream
                    = GetSectorChain(RootEntry.StartSector, SectorType.Normal);

                var miniStreamView
                    = new StreamView(miniStream, GetSectorSize(), RootStorage.Size, SourceStream);

                long ptr = 0;

                var nMinisectors = (int) (miniStreamView.Length/Sector.MinisectorSize);

                while (ptr < nMinisectors)
                {
                    //AssureLength(miniStreamView, (int)miniFATView.Length);

                    var id = miniFATView.ReadInt32();
                    ptr += 4;

                    if (id != Sector.FreeSector) continue;
                    var miniSector = new Sector(Sector.MinisectorSize, SourceStream)
                    {
                        Id = (int) ((ptr - 4)/4),
                        Type = SectorType.Mini
                    };

                    miniStreamView.Seek(miniSector.Id*Sector.MinisectorSize, SeekOrigin.Begin);
                    miniStreamView.Read(miniSector.GetData(), 0, Sector.MinisectorSize);

                    freeList.Enqueue(miniSector);
                }
            }

            return freeList;
        }
        #endregion

        #region SetData
        /// <summary>
        ///     Sets the data for the current stream
        /// </summary>
        /// <param name="cfItem"></param>
        /// <param name="buffer"></param>
        internal void SetData(CFItem cfItem, Byte[] buffer)
        {
            SetStreamData(cfItem, buffer);
        }

        /// <summary>
        ///     Sets the data for the current stream
        /// </summary>
        /// <param name="cfItem"></param>
        /// <param name="buffer"></param>
        /// <exception cref="CFException">Raised when <see cref="buffer" /> is null</exception>
        private void SetStreamData(CFItem cfItem, Byte[] buffer)
        {
            if (buffer == null)
                throw new CFException("Parameter [buffer] cannot be null");

            // Quick and dirty :-)
            if (buffer.Length == 0) return;

            var directoryEntry = cfItem.DirEntry;

            var sectorType = SectorType.Normal;
            var sectorSize = GetSectorSize();

            if (buffer.Length < _header.MinSizeStandardStream)
            {
                sectorType = SectorType.Mini;
                sectorSize = Sector.MinisectorSize;
            }

            // Check for transition ministream -> stream:
            // Only in this case we need to free old sectors,
            // otherwise they will be overwritten.

            if (directoryEntry.StartSector != Sector.Endofchain)
            {
                if (
                    (buffer.Length < _header.MinSizeStandardStream &&
                     directoryEntry.Size > _header.MinSizeStandardStream)
                    ||
                    (buffer.Length > _header.MinSizeStandardStream &&
                     directoryEntry.Size < _header.MinSizeStandardStream)
                    )
                {
                    if (directoryEntry.Size < _header.MinSizeStandardStream)
                    {
                        FreeMiniChain(GetMiniSectorChain(directoryEntry.StartSector), _eraseFreeSectors);
                    }
                    else
                    {
                        FreeChain(GetNormalSectorChain(directoryEntry.StartSector), _eraseFreeSectors);
                    }

                    directoryEntry.Size = 0;
                    directoryEntry.StartSector = Sector.Endofchain;
                }
            }

            var sectorChain
                = GetSectorChain(directoryEntry.StartSector, sectorType);

            Queue<Sector> freeList = null;

            if (_sectorRecycle)
                freeList = FindFreeSectors(sectorType); // Collect available free sectors

            var streamView = new StreamView(sectorChain, sectorSize, buffer.Length, freeList, SourceStream);

            streamView.Write(buffer, 0, buffer.Length);

            switch (sectorType)
            {
                case SectorType.Normal:
                    SetNormalSectorChain(streamView.BaseSectorChain);
                    break;

                case SectorType.Mini:
                    SetMiniSectorChain(streamView.BaseSectorChain);
                    break;
            }

            if (streamView.BaseSectorChain.Count > 0)
            {
                directoryEntry.StartSector = streamView.BaseSectorChain[0].Id;
                directoryEntry.Size = buffer.Length;
            }
            else
            {
                directoryEntry.StartSector = Sector.Endofchain;
                directoryEntry.Size = 0;
            }
        }
        #endregion

        #region GetData
        /// <summary>
        ///     Gets data from the <see cref="cFStream" />
        /// </summary>
        /// <param name="cFStream"></param>
        /// <param name="offset"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        /// <exception cref="CFDisposedException">Raised when the file is closed</exception>
        internal byte[] GetData(CFStream cFStream, long offset, ref int count)
        {
            var directoryEntry = cFStream.DirEntry;
            count = (int) Math.Min(directoryEntry.Size - offset, count);

            StreamView streamView;

            if (directoryEntry.Size < _header.MinSizeStandardStream)
            {
                streamView
                    = new StreamView(GetSectorChain(directoryEntry.StartSector, SectorType.Mini), Sector.MinisectorSize,
                        directoryEntry.Size,
                        SourceStream);
            }
            else
            {
                streamView = new StreamView(GetSectorChain(directoryEntry.StartSector, SectorType.Normal),
                    GetSectorSize(), directoryEntry.Size,
                    SourceStream);
            }

            var result = new byte[count];
            streamView.Seek(offset, SeekOrigin.Begin);
            streamView.Read(result, 0, result.Length);

            return result;
        }

        /// <summary>
        ///     Gets data from the <see cref="cFStream" />
        /// </summary>
        /// <param name="cFStream"></param>
        /// <returns></returns>
        /// <exception cref="CFDisposedException">Raised when the file is closed</exception>
        internal byte[] GetData(CFStream cFStream)
        {
            if (IsClosed)
                throw new CFDisposedException("Compound File closed: cannot access data");

            byte[] result;

            var directoryEntry = cFStream.DirEntry;

            if (directoryEntry.Size < _header.MinSizeStandardStream)
            {
                var miniView
                    = new StreamView(GetSectorChain(directoryEntry.StartSector, SectorType.Mini), Sector.MinisectorSize,
                        directoryEntry.Size,
                        SourceStream);

                var br = new BinaryReader(miniView);

                result = br.ReadBytes((int) directoryEntry.Size);
                br.Close();
            }
            else
            {
                var sView
                    = new StreamView(GetSectorChain(directoryEntry.StartSector, SectorType.Normal), GetSectorSize(),
                        directoryEntry.Size,
                        SourceStream);

                result = new byte[(int) directoryEntry.Size];

                sView.Read(result, 0, result.Length);
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

        #region LowSaturation
        private static int LowSaturation(int i)
        {
            return i > 0 ? i : 0;
        }
        #endregion

        #region IDisposable Members
        void IDisposable.Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

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

                            RootStorage = null; // Some problem releasing resources...
                            _header = null;
                            _directoryEntries.Clear();
                            _directoryEntries = null;
                            _lockObject = null;
                            _buffer = null;
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
        #endregion

        #region GetAllNamedEntries
        /// <summary>
        ///     Get a list of all entries which start with the given <see cref="entryName" />
        /// </summary>
        /// <param name="entryName">Name of entries to retrive</param>
        /// <param name="parentSibling">
        ///     The parent id from the node where you want to find the named entries,
        ///     use null if you want to search in all nodes
        /// </param>
        /// <returns>A list of name-matching entries</returns>
        /// <remarks>
        ///     This function is aimed to speed up entity lookup in
        ///     flat-structure files (only one or little more known entries)
        ///     without the performance penalty related to entities hierarchy constraints.
        ///     There is no implied hierarchy in the returned list.
        /// </remarks>
        public IList<CFItem> GetAllNamedEntries(string entryName, int? parentSibling)
        {
            var result = new List<CFItem>();

            foreach (var directoryEntry in _directoryEntries)
            {
                if (directoryEntry.StgType == StgType.StgInvalid || !directoryEntry.GetEntryName().StartsWith(entryName))
                    continue;
                if (directoryEntry.LeftSibling != parentSibling && parentSibling != null) continue;
                var cfItem = directoryEntry.StgType == StgType.StgStorage
                    ? new CFStorage(this, directoryEntry)
                    : (CFItem) new CFStream(this, directoryEntry);

                result.Add(cfItem);
            }

            return result;
        }
        #endregion
    }
}