using System;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Interfaces
{
    /// <summary>
    /// The compound file storage interface
    /// </summary>
    public interface ICFStorage : ICFItem
    {
        /// <summary>
        ///     Gets a named
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
        ICFStream GetStream(String streamName);

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
        bool ExistsStream(string streamName);

        /// <summary>
        ///     Gets a named storage contained in the current one if existing.
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
        ICFStorage GetStorage(String storageName);

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
        bool ExistsStorage(string storageName);

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
        void VisitEntries(VisitedEntryAction action, bool recursive);

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
        ///     var castList = (List<string />)args[0];
        ///     castList.Add(item.Name);
        /// };
        /// 
        /// var list = new List<string />();
        /// 
        /// cf.RootStorage.VisitEntries(va, true, list);
        /// 
        /// list.ForEach(tw.WriteLine);
        /// 
        /// tw.Close();
        /// </code>
        /// </example>
        void VisitEntries(VisitedEntryParamsAction action, bool recursive, params object[] args);
    }
}