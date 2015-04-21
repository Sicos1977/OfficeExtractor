using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using CompoundFileStorage;
using CompoundFileStorage.Exceptions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentServices.Modules.Extractors.OfficeExtractor.Biff8;
using DocumentServices.Modules.Extractors.OfficeExtractor.Exceptions;
using DocumentServices.Modules.Extractors.OfficeExtractor.Helpers;

namespace DocumentServices.Modules.Extractors.OfficeExtractor
{
    /// <summary>
    /// This class is used as a placeholder for all Excel related methods
    /// </summary>
    internal static class Excel
    {
        #region SaveToFolder
        /// <summary>
        /// This method saves all the Excel embedded binary objects from the <paramref name="inputFile"/> to the
        /// <see cref="outputFolder"/>
        /// </summary>
        /// <param name="inputFile">The binary Excel file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <param name="storageName">The complete or part of the name from the storage that needs to be saved</param>
        /// <returns></returns>
        /// <exception cref="OEFileIsPasswordProtected">Raised when the <paramref name="inputFile"/> is password protected</exception>
        /// <exception cref="OEFileIsCorrupt">Raised when the file is corrupt</exception>
        public static List<string> SaveToFolder(string inputFile, string outputFolder, string storageName)
        {
            using (var compoundFile = new CompoundFile(inputFile))
            {
                try
                {
                    if (IsPasswordProtected(compoundFile))
                        throw new OEFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                            "' is password protected");
                }
                catch (CFCorruptedFileException)
                {
                    throw new OEFileIsCorrupt("The file '" + Path.GetFileName(inputFile) + "' is corrupt");
                }

                var result = new List<string>();

                foreach (var child in compoundFile.RootStorage.Children)
                {
                    var childStorage = child as CFStorage;
                    if (childStorage == null) continue;
                    if (!childStorage.Name.StartsWith(storageName)) continue;

                    var extractedFileName = Extraction.SaveFromStorageNode(childStorage, outputFolder);
                    if (extractedFileName != null)
                        result.Add(extractedFileName);
                }

                return result;
            }
        }
        #endregion   

        #region IsPasswordProtected
        /// <summary>
        /// Returns true when the Excel file is password protected
        /// </summary>
        /// <param name="compoundFile">The Excel file to check</param>
        /// <returns></returns>
        /// <exception cref="OEFileIsCorrupt">Raised when the file is corrupt</exception>
        public static bool IsPasswordProtected(CompoundFile compoundFile)
        {
            try
            {
                if (compoundFile.RootStorage.ExistsStream("EncryptedPackage")) return true;
                if (!compoundFile.RootStorage.ExistsStream("WorkBook"))
                    throw new OEFileIsCorrupt("Could not find the WorkBook stream in the file '" +
                                              compoundFile.FileName + "'");

                var stream = compoundFile.RootStorage.GetStream("WorkBook") as CFStream;
                if (stream == null) return false;

                var bytes = stream.GetData();
                using (var memoryStream = new MemoryStream(bytes))
                using (var binaryReader = new BinaryReader(memoryStream))
                {
                    // Get the record type, at the beginning of the stream this should always be the BOF
                    var recordType = binaryReader.ReadUInt16();

                    // Something seems to be wrong, we would expect a BOF but for some reason it isn't so stop it
                    if (recordType != 0x809)
                        throw new OEFileIsCorrupt("The file '" + Path.GetFileName(compoundFile.FileName) +
                                                  "' is corrupt");

                    var recordLength = binaryReader.ReadUInt16();
                    binaryReader.BaseStream.Position += recordLength;

                    // Search after the BOF for the FilePass record, this starts with 2F hex
                    recordType = binaryReader.ReadUInt16();
                    if (recordType != 0x2F) return false;
                    binaryReader.ReadUInt16();
                    var filePassRecord = new FilePassRecord(memoryStream);
                    var key = Biff8EncryptionKey.Create(filePassRecord.DocId);
                    return !key.Validate(filePassRecord.SaltData, filePassRecord.SaltHash);
                }
            }
            catch (OEExcelConfiguration)
            {
                // If we get an OCExcelConfiguration exception it means we have an unknown encryption
                // type so we return a false so that Excel itself can figure out if the file is password
                // protected
                return false;
            }
            catch (CFFileFormatException)
            {
                // It seems the file is just a normal Microsoft Office 2007 and up Open XML file
                return false;
            }
        }
        #endregion

        #region SetWorkbookVisibility
        /// <summary>
        /// When a Excel document is embedded in for example a Word document the Workbook
        /// is set to hidden. Don't know why Microsoft does this but they do. To solve this
        /// problem we seek the WINDOW1 record in the BOF record of the stream. In there a
        /// gbit structure is located. The first bit in this structure controls the visibility
        /// of the workbook, so we check if this bit is set to 1 (hidden) en is so set it to 0.
        /// Normally a Workbook stream only contains one WINDOW record but when it is embedded
        /// it will contain 2 or more records.
        /// </summary>
        /// <param name="compoundFile"></param>
        /// <exception cref="OEFileIsCorrupt">Raised when the <paramref name="compoundFile"/> does not have a Workbook stream</exception>
        public static void SetWorkbookVisibility(CompoundFile compoundFile)
        {
            if (!compoundFile.RootStorage.ExistsStream("WorkBook"))
                throw new OEFileIsCorrupt("Could not check workbook visibility because the WorkBook stream is not present");

            try
            {
                var stream = compoundFile.RootStorage.GetStream("WorkBook") as CFStream;
                if (stream == null) return;

                var bytes = stream.GetData();

                using (var memoryStream = new MemoryStream(bytes))
                using (var binaryReader = new BinaryReader(memoryStream))
                {
                    // Get the record type, at the beginning of the stream this should always be the BOF
                    var recordType = binaryReader.ReadUInt16();
                    var recordLength = binaryReader.ReadUInt16();

                    // Something seems to be wrong, we would expect a BOF but for some reason it isn't 
                    if (recordType != 0x809)
                        throw new OEFileIsCorrupt("The file is corrupt");

                    binaryReader.BaseStream.Position += recordLength;

                    while (binaryReader.BaseStream.Position < binaryReader.BaseStream.Length)
                    {
                        recordType = binaryReader.ReadUInt16();
                        recordLength = binaryReader.ReadUInt16();

                        // Window1 record (0x3D)
                        if (recordType == 0x3D)
                        {
                            // ReSharper disable UnusedVariable
                            var xWn = binaryReader.ReadUInt16();
                            var yWn = binaryReader.ReadUInt16();
                            var dxWn = binaryReader.ReadUInt16();
                            var dyWn = binaryReader.ReadUInt16();
                            // ReSharper restore UnusedVariable

                            // The grbit contains the bit that hides the sheet
                            var grbit = binaryReader.ReadBytes(2);
                            var bitArray = new BitArray(grbit);

                            // When the bit is set then unset it (bitArray.Get(0) == true)
                            if (bitArray.Get(0))
                            {
                                bitArray.Set(0, false);

                                // Copy the byte back into the stream, 2 positions back so that we overwrite the old bytes
                                bitArray.CopyTo(bytes, (int)binaryReader.BaseStream.Position - 2);
                            }

                            break;
                        }
                        binaryReader.BaseStream.Position += recordLength;
                    }
                }

                stream.SetData(bytes);
            }
            catch (Exception exception)
            {
                throw new OEFileIsCorrupt(
                    "Could not check workbook visibility because the file seems to be corrupt", exception);
            }
            
        }

        /// <summary>
        /// This method sets the workbook in an Open XML Format Excel file to visible
        /// </summary>
        /// <param name="spreadSheetDocument">The Open XML Format Excel file as a memorystream</param>
        /// <exception cref="OEFileIsCorrupt">Raised when the <paramref name="spreadSheetDocument"/> is corrupt</exception>
        public static MemoryStream SetWorkbookVisibility(MemoryStream spreadSheetDocument)
        {
            try
            {
                using (var spreadsheetDocument = SpreadsheetDocument.Open(spreadSheetDocument, true))
                {
                    var bookViews = spreadsheetDocument.WorkbookPart.Workbook.BookViews;
                    foreach (var bookView in bookViews)
                    {
                        var workBookView = (WorkbookView)bookView;
                        if (workBookView.Visibility.Value == VisibilityValues.Hidden ||
                            workBookView.Visibility.Value == VisibilityValues.VeryHidden)
                            workBookView.Visibility.Value = VisibilityValues.Visible;
                    }

                    spreadsheetDocument.WorkbookPart.Workbook.Save();
                }

                return spreadSheetDocument;
            }
            catch (Exception exception)
            {
                throw new OEFileIsCorrupt(
                    "Could not check workbook visibility because the file seems to be corrupt", exception);
            }
        }
        #endregion
    }
}
