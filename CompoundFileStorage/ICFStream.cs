using System;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{
    public interface ICFStream : ICFItem
    {
        /// <summary>
        /// Get the data associated with the stream object.
        /// </summary>
        /// <example>
        /// <code>
        ///     CompoundFile cf2 = new CompoundFile("AFileName.cfs");
        ///     CFStream st = cf2.RootStorage.GetStream("MyStream");
        ///     byte[] buffer = st.GetData();
        /// </code>
        /// </example>
        /// <returns>Array of byte containing stream data</returns>
        /// <exception cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFDisposedException">
        /// Raised when the owner compound file has been closed.
        /// </exception>
        Byte[] GetData();

        /// <summary>
        /// Get <paramref name="count"/> bytes associated with the stream object, starting from
        /// a provided <paramref name="offset"/>. When method returns, count will contain the
        /// effective count of bytes read.
        /// </summary>
        /// <example>
        /// <code>
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
        /// Raised when the owner compound file has been closed.
        /// </exception>
        Byte[] GetData(long offset, ref int count);
    }
}