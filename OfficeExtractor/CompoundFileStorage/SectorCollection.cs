using System;
using System.Collections;
using System.Collections.Generic;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{
    /// <summary>
    ///     Action to implement when transaction support - sector has to be written to the underlying stream (see specs).
    /// </summary>
    public delegate void Ver3SizeLimitReached();

    /// <summary>
    ///     Ad-hoc Heap Friendly sector collection to avoid using large array that may create some problem to GC collection
    ///     (see http://www.simple-talk.com/dotnet/.net-framework/the-dangers-of-the-large-object-heap/ )
    /// </summary>
    internal class SectorCollection : IList<Sector>
    {
        #region Fields
        /// <summary>
        ///     0x7FFFFF00 for Version 4
        /// </summary>
        private const int MaxSectorV4CountLockRange = 524287;

        private const int SliceSize = 4096;
        private readonly List<ArrayList> _largeArraySlices = new List<ArrayList>();
        private bool _sizeLimitReached;
        #endregion

        #region Events
        public event Ver3SizeLimitReached OnVer3SizeLimitReached;
        #endregion

        #region DoCheckSizeLimitReached
        private void DoCheckSizeLimitReached()
        {
            if (_sizeLimitReached || (Count - 1 <= MaxSectorV4CountLockRange)) return;
            if (OnVer3SizeLimitReached != null)
                OnVer3SizeLimitReached();

            _sizeLimitReached = true;
        }
        #endregion

        #region IList<T> Members
        public int IndexOf(Sector item)
        {
            throw new NotImplementedException();
        }

        public void Insert(int index, Sector item)
        {
            throw new NotImplementedException();
        }

        public void RemoveAt(int index)
        {
            throw new NotImplementedException();
        }

        public Sector this[int index]
        {
            get
            {
                var itemIndex = index/SliceSize;
                var itemOffset = index%SliceSize;

                if ((index > -1) && (index < Count))
                    return (Sector) _largeArraySlices[itemIndex][itemOffset];

                throw new ArgumentOutOfRangeException("index", index, "Argument out of range");
            }

            set
            {
                var itemIndex = index/SliceSize;
                var itemOffset = index%SliceSize;

                if (index > -1 && index < Count)
                {
                    _largeArraySlices[itemIndex][itemOffset] = value;
                }
                else
                    throw new ArgumentOutOfRangeException("index", index, "Argument out of range");
            }
        }
        #endregion

        #region ICollection<T> Members
        public void Add(Sector item)
        {
            DoCheckSizeLimitReached();

            var itemIndex = Count/SliceSize;

            if (itemIndex < _largeArraySlices.Count)
            {
                _largeArraySlices[itemIndex].Add(item);
                Count++;
            }
            else
            {
                var ar = new ArrayList(SliceSize) {item};
                _largeArraySlices.Add(ar);
                Count++;
            }
        }

        public void Clear()
        {
            foreach (var slice in _largeArraySlices)
            {
                slice.Clear();
            }

            _largeArraySlices.Clear();

            Count = 0;
        }

        public bool Contains(Sector item)
        {
            throw new NotImplementedException();
        }

        public void CopyTo(Sector[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public int Count { get; private set; }

        public bool IsReadOnly
        {
            get { return false; }
        }

        public bool Remove(Sector item)
        {
            throw new NotImplementedException();
        }
        #endregion

        #region IEnumerable<T> Members
        public IEnumerator<Sector> GetEnumerator()
        {
            foreach (var largeArraySlice in _largeArraySlices)
            {
                foreach (var t in largeArraySlice)
                    yield return (Sector) t;
            }
        }
        #endregion

        #region IEnumerable Members
        IEnumerator IEnumerable.GetEnumerator()
        {
            foreach (var largeArraySlice in _largeArraySlices)
            {
                for (var j = 0; j < largeArraySlice.Count; j++)
                    yield return largeArraySlice[j];
            }
        }
        #endregion
    }
}