using System;
using System.IO;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{
    internal enum SectorType
    {
        Normal,
        Mini,
        FAT,
        DIFAT,
        RangeLockSector,
        Directory
    }

    internal class Sector : IDisposable
    {
        public const int FREESECT = unchecked((int) 0xFFFFFFFF);
        public const int ENDOFCHAIN = unchecked((int) 0xFFFFFFFE);
        public const int FATSECT = unchecked((int) 0xFFFFFFFD);
        public const int DIFSECT = unchecked((int) 0xFFFFFFFC);
        public static int MINISECTOR_SIZE = 64;
        private readonly object lockObject = new Object();
        private readonly Stream stream;
        private byte[] data;

        private bool dirtyFlag;
        private int id = -1;

        private int size;


        public Sector(int size, Stream stream)
        {
            this.size = size;
            this.stream = stream;
        }

        public Sector(int size, byte[] data)
        {
            this.size = size;
            this.data = data;
            stream = null;
        }

        public Sector(int size)
        {
            this.size = size;
            data = null;
            stream = null;
        }

        public bool DirtyFlag
        {
            get { return dirtyFlag; }
            set { dirtyFlag = value; }
        }

        public bool IsStreamed
        {
            get { return (stream != null && size != MINISECTOR_SIZE) && (id*size) + size < stream.Length; }
        }

        internal SectorType Type { get; set; }

        public int Id
        {
            get { return id; }
            set { id = value; }
        }

        public int Size
        {
            get { return size; }
        }

        public byte[] GetData()
        {
            if (data == null)
            {
                data = new byte[size];

                if (IsStreamed)
                {
                    stream.Seek(size + id*(long) size, SeekOrigin.Begin);
                    stream.Read(data, 0, size);
                }
            }

            return data;
        }

        //public void SetSectorData(byte[] b)
        //{
        //    this.data = b;
        //}

        //public void FillData(byte b)
        //{
        //    if (data != null)
        //    {
        //        for (int i = 0; i < data.Length; i++)
        //        {
        //            data[i] = b;
        //        }
        //    }
        //}

        public void ZeroData()
        {
            data = new byte[size];
            dirtyFlag = true;
        }

        internal void ReleaseData()
        {
            data = null;
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
                if (!_disposed)
                {
                    lock (lockObject)
                    {
                        if (disposing)
                        {
                            // Call from user code...
                        }

                        data = null;
                        dirtyFlag = false;
                        id = Sector.ENDOFCHAIN;
                        size = 0;
                    }
                }
            }
            finally
            {
                _disposed = true;
            }
        }

        #region IDisposable Members
        private bool _disposed; //false

        void IDisposable.Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}