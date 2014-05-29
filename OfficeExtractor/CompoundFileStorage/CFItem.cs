using System;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Exceptions;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Interfaces;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{
    /// <summary>
    ///     Abstract base class for Structured Storage entities.
    /// </summary>
    public abstract class CFItem : IComparable, ICFItem
    {
        #region Fields
        private readonly CompoundFile _compoundFile;
        #endregion

        #region Properties
        protected CompoundFile CompoundFile
        {
            get { return _compoundFile; }
        }

        /// <summary>
        ///     Get entity name
        /// </summary>
        public string Name
        {
            get
            {
                var name = DirEntry.GetEntryName();
                return !string.IsNullOrEmpty(name) ? name.TrimEnd('\0') : string.Empty;
            }
        }

        /// <summary>
        ///     Size in bytes of the item. It has a valid value
        ///     only if entity is a stream, otherwise it is setted to zero.
        /// </summary>
        public long Size
        {
            get { return DirEntry.Size; }
        }


        /// <summary>
        ///     Returns true if item is Storage
        /// </summary>
        /// <remarks>
        ///     This check doesn't use reflection or runtime type information
        ///     and doesn't suffer related performance penalties.
        /// </remarks>
        public bool IsStorage
        {
            get { return DirEntry.StgType == StgType.StgStorage; }
        }

        /// <summary>
        ///     Returns true if item is a Stream
        /// </summary>
        /// <remarks>
        ///     This check doesn't use reflection or runtime type information
        ///     and doesn't suffer related performance penalties.
        /// </remarks>
        public bool IsStream
        {
            get { return DirEntry.StgType == StgType.StgStream; }
        }

        /// <summary>
        ///     Returnstrue if item is the Root Storage
        /// </summary>
        /// <remarks>
        ///     This check doesn't use reflection or runtime type information
        ///     and doesn't suffer related performance penalties.
        /// </remarks>
        public bool IsRoot
        {
            get { return DirEntry.StgType == StgType.StgRoot; }
        }

        /// <summary>
        ///     Get/Set the Creation Date of the current item
        /// </summary>
        public DateTime CreationDate
        {
            get { return DateTime.FromFileTime(BitConverter.ToInt64(DirEntry.CreationDate, 0)); }

            set
            {
                if (DirEntry.StgType != StgType.StgStream && DirEntry.StgType != StgType.StgRoot)
                    DirEntry.CreationDate = BitConverter.GetBytes((value.ToFileTime()));
                else
                    throw new CFException("CreationDate can only be set on storage entries");
            }
        }

        /// <summary>
        ///     Get/Set the Modify Date of the current item
        /// </summary>
        public DateTime ModifyDate
        {
            get { return DateTime.FromFileTime(BitConverter.ToInt64(DirEntry.ModifyDate, 0)); }

            set
            {
                if (DirEntry.StgType != StgType.StgStream && DirEntry.StgType != StgType.StgRoot)
                    DirEntry.ModifyDate = BitConverter.GetBytes((value.ToFileTime()));
                else
                    throw new CFException("ModifyDate can only be set on storage entries");
            }
        }

        /// <summary>
        ///     Get/Set Object class Guid for Root and Storage entries.
        /// </summary>
        public Guid CLSID
        {
            get { return DirEntry.StorageCLSID; }
            set
            {
                if (DirEntry.StgType != StgType.StgStream)
                    DirEntry.StorageCLSID = value;
                else
                    throw new CFException("Object class GUID can only be set on Root and Storage entries");
            }
        }
        #endregion

        #region CheckDisposed
        protected void CheckDisposed()
        {
            if (_compoundFile.IsClosed)
                throw new CFDisposedException(
                    "The owner compound file has been closed and owned items have been invalidated");
        }
        #endregion

        #region CFItem
        protected CFItem()
        {
        }

        protected CFItem(CompoundFile compoundFile)
        {
            _compoundFile = compoundFile;
        }
        #endregion

        #region IDirectoryEntry Members
        internal IDirectoryEntry DirEntry { get; set; }

        internal int CompareTo(CFItem other)
        {
            return DirEntry.CompareTo(other.DirEntry);
        }
        #endregion

        #region IComparable Members
        public int CompareTo(object obj)
        {
            return DirEntry.CompareTo(((CFItem) obj).DirEntry);
        }
        #endregion

        #region Operators
        public static bool operator ==(CFItem leftItem, CFItem rightItem)
        {
            // If both are null, or both are same instance, return true.
            if (ReferenceEquals(leftItem, rightItem))
                return true;

            // If one is null, but not both, return false.
            if (((object) leftItem == null) || ((object) rightItem == null))
                return false;

            // Return true if the fields match:
            return leftItem.CompareTo(rightItem) == 0;
        }

        public static bool operator !=(CFItem leftItem, CFItem rightItem)
        {
            return !(leftItem == rightItem);
        }
        #endregion

        #region Equals
        public override bool Equals(object obj)
        {
            return CompareTo(obj) == 0;
        }
        #endregion

        #region GetHashCode
        public override int GetHashCode()
        {
            return DirEntry.GetEntryName().GetHashCode();
        }
        #endregion
    }
}