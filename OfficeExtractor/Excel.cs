using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;
using OpenMcdf;

/*
   Copyright 2013 - 2016 Kees van Spelde

   Licensed under The Code Project Open License (CPOL) 1.02;
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.codeproject.com/info/cpol10.aspx

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

namespace OfficeExtractor
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
        /// <returns></returns>
        /// <exception cref="OEFileIsPasswordProtected">Raised when the <paramref name="inputFile"/> is password protected</exception>
        /// <exception cref="OEFileIsCorrupt">Raised when the file is corrupt</exception>
        public static List<string> SaveToFolder(string inputFile, string outputFolder)
        {
            var fileName = Path.GetFileName(inputFile);

            using (var compoundFile = new CompoundFile(inputFile))
            {
                if (IsPasswordProtected(compoundFile, fileName))
                    throw new OEFileIsPasswordProtected("The file '" + fileName + "' is password protected");

                var result = new List<string>();
                Action<CFItem> entries = storage =>
                {
                    var childStorage = storage as CFStorage;
                    if (childStorage == null || !childStorage.Name.StartsWith("MBD")) return;
                    var extractedFileName = Extraction.SaveFromStorageNode(childStorage, outputFolder);
                    if (extractedFileName != null)
                        result.Add(extractedFileName);
                };

                compoundFile.RootStorage.VisitEntries(entries, false);
                return result;
            }
        }
        #endregion   

        #region IsPasswordProtected
        /// <summary>
        /// Returns true when the Excel file is password protected
        /// </summary>
        /// <param name="compoundFile">The Excel file to check</param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        /// <exception cref="OEFileIsCorrupt">Raised when the file is corrupt</exception>
        public static bool IsPasswordProtected(CompoundFile compoundFile, string fileName)
        {
            try
            {
                if (compoundFile.RootStorage.TryGetStream("EncryptedPackage") != null) return true;

                var stream = compoundFile.RootStorage.TryGetStream("WorkBook");
                if (stream == null)
                    compoundFile.RootStorage.TryGetStream("Book");

                if (stream == null)
                    throw new OEFileIsCorrupt("Could not find the WorkBook or Book stream in the file '" + fileName + "'");

                var bytes = stream.GetData();
                using (var memoryStream = new MemoryStream(bytes))
                using (var binaryReader = new BinaryReader(memoryStream))
                {
                    // Get the record type, at the beginning of the stream this should always be the BOF
                    var recordType = binaryReader.ReadUInt16();

                    // Something seems to be wrong, we would expect a BOF but for some reason it isn't so stop it
                    if (recordType != 0x809)
                        throw new OEFileIsCorrupt("The file '" + fileName + "' is corrupt");

                    var recordLength = binaryReader.ReadUInt16();
                    binaryReader.BaseStream.Position += recordLength;

                    // Search after the BOF for the FilePass record, this starts with 2F hex
                    recordType = binaryReader.ReadUInt16();
                    return recordType == 0x2F;
                }
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
        /// <param name="rootStorage">The <see cref="CFStorage">Root storage</see> of a <see cref="CompoundFile"/></param>
        /// <exception cref="OEFileIsCorrupt">Raised when the <paramref name="rootStorage"/> does not have a Workbook stream</exception>
        public static void SetWorkbookVisibility(CFStorage rootStorage)
        {
            var stream = rootStorage.TryGetStream("WorkBook");
            if (stream == null)
                throw new OEFileIsCorrupt("Could not check workbook visibility because the WorkBook stream is not present");

            try
            {
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
