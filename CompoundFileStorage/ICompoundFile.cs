using System;
using System.Collections.Generic;
using System.IO;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{
    public interface ICompoundFile
    {
        bool ValidationExceptionEnabled { get; }

        /// <summary>
        /// Return true if this compound file has been 
        /// loaded from an existing file or stream
        /// </summary>
        bool HasSourceStream { get; }

        /// <summary>
        /// The entry point object that represents the 
        /// root of the structures tree to get or set storage or
        /// stream data.
        /// </summary>
        /// <example>
        /// <code>
        /// 
        ///    //Create a compound file
        ///    string FILENAME = "MyFileName.cfs";
        ///    CompoundFile ncf = new CompoundFile();
        ///
        ///    CFStorage l1 = ncf.RootStorage.AddStorage("Storage Level 1");
        ///
        ///    l1.AddStream("l1ns1");
        ///    l1.AddStream("l1ns2");
        ///    l1.AddStream("l1ns3");
        ///    CFStorage l2 = l1.AddStorage("Storage Level 2");
        ///    l2.AddStream("l2ns1");
        ///    l2.AddStream("l2ns2");
        ///
        ///    ncf.Save(FILENAME);
        ///    ncf.Close();
        /// </code>
        /// </example>
        ICFStorage RootStorage { get; }

        CFSVersion Version { get; }

        /// <summary>
        /// Commit data changes since the previously commit operation
        /// to the underlying supporting stream or file on the disk.
        /// </summary>
        /// <remarks>
        /// This method can be used
        /// only if the supporting stream has been opened in 
        /// <see cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.UpdateMode">Update mode</see>.
        /// </remarks>
        void Commit();

        /// <summary>
        /// Commit data changes since the previously commit operation
        /// to the underlying supporting stream or file on the disk.
        /// </summary>
        /// <param name="releaseMemory">If true, release loaded sectors to limit memory usage but reduces following read operations performance</param>
        /// <remarks>
        /// This method can be used only if 
        /// the supporting stream has been opened in 
        /// <see cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.UpdateMode">Update mode</see>.
        /// </remarks>
        void Commit(bool releaseMemory);

        /// <summary>
        /// Saves the in-memory image of Compound File to a file.
        /// </summary>
        /// <param name="fileName">File name to write the compound file to</param>
        /// <exception cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFException">Raised if destination file is not seekable</exception>
        void Save(String fileName);

        /// <summary>
        /// Saves the in-memory image of Compound File to a stream.
        /// </summary>        
        /// <remarks>
        /// Destination Stream must be seekable.
        /// </remarks>
        /// <param name="stream">The stream to save compound File to</param>
        /// <exception cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFException">Raised if destination stream is not seekable</exception>
        /// <exception cref="T:DocumentServices.Modules.Extractors.OfficeExtractor.OLECompoundFileStorage.CFDisposedException">Raised if Compound File Storage has been already disposed</exception>
        /// <example>
        /// <code>
        ///    MemoryStream ms = new MemoryStream(size);
        ///
        ///    CompoundFile cf = new CompoundFile();
        ///    CFStorage st = cf.RootStorage.AddStorage("MyStorage");
        ///    CFStream sm = st.AddStream("MyStream");
        ///
        ///    byte[] b = new byte[]{0x00,0x01,0x02,0x03};
        ///
        ///    sm.SetData(b);
        ///    cf.Save(ms);
        ///    cf.Close();
        /// </code>
        /// </example>
        void Save(Stream stream);

        /// <summary>
        /// Close the Compound File object <see cref="T:OpenMcdf.CompoundFile">CompoundFile</see> and
        /// free all associated resources (e.g. open file handle and allocated memory).
        /// <remarks>
        /// When the <see cref="T:OpenMcdf.CompoundFile.Close()">Close</see> method is called,
        /// all the associated stream and storage objects are invalidated:
        /// any operation invoked on them will produce a <see cref="T:OpenMcdf.CFDisposedException">CFDisposedException</see>.
        /// </remarks>
        /// </summary>
        /// <example>
        /// <code>
        ///    const String FILENAME = "CompoundFile.cfs";
        ///    CompoundFile cf = new CompoundFile(FILENAME);
        ///
        ///    CFStorage st = cf.RootStorage.GetStorage("MyStorage");
        ///    cf.Close();
        ///
        ///    try
        ///    {
        ///        byte[] temp = st.GetStream("MyStream").GetData();
        ///        
        ///        // The following line will fail because back-end object has been closed
        ///        Assert.Fail("Stream without media");
        ///    }
        ///    catch (Exception ex)
        ///    {
        ///        Assert.IsTrue(ex is CFDisposedException);
        ///    }
        /// </code>
        /// </example>
        void Close();

        /// <summary>
        /// Get a list of all entries with a given name contained in the document.
        /// </summary>
        /// <param name="entryName">Name of entries to retrive</param>
        /// <returns>A list of name-matching entries</returns>
        /// <remarks>This function is aimed to speed up entity lookup in 
        /// flat-structure files (only one or little more known entries)
        /// without the performance penalty related to entities hierarchy constraints.
        /// There is no implied hierarchy in the returned list.
        /// </remarks>
        IList<ICFItem> GetAllNamedEntries(String entryName);
    }
}