using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;
using OpenMcdf;

//
// Excel.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2020 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

namespace OfficeExtractor
{
    /// <summary>
    /// This class is used as a placeholder for all Excel related methods
    /// </summary>
    internal class Excel
    {
        #region Fields
        /// <summary>
        ///     <see cref="Extraction"/>
        /// </summary>
        private Extraction _extraction;
        #endregion

        #region Properties
        /// <summary>
        /// Returns a reference to the Extraction class when it already exists or creates a new one
        /// when it doesn't
        /// </summary>
        private Extraction Extraction
        {
            get
            {
                if (_extraction != null)
                    return _extraction;

                _extraction = new Extraction();
                return _extraction;
            }
        }
        #endregion

        #region Extract
        /// <summary>
        /// This method saves all the Excel embedded binary objects from the <paramref name="inputFile"/> to the
        /// <paramref name="outputFolder"/>
        /// </summary>
        /// <param name="inputFile">The binary Excel file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns></returns>
        /// <exception cref="OEFileIsPasswordProtected">Raised when the <paramref name="inputFile"/> is password protected</exception>
        /// <exception cref="OEFileIsCorrupt">Raised when the file is corrupt</exception>
        internal List<string> Extract(string inputFile, string outputFolder)
        {
            using (var compoundFile = new CompoundFile(inputFile))
            {
                var result = new List<string>();

                void Entries(CFItem storage)
                {
                    var childStorage = storage as CFStorage;
                    if (childStorage == null || !childStorage.Name.StartsWith("MBD")) return;
                    var extractedFileName = Extraction.SaveFromStorageNode(childStorage, outputFolder);
                    if (extractedFileName != null) result.Add(extractedFileName);
                }

                compoundFile.RootStorage.VisitEntries(Entries, false);
                return result;
            }
        }
        #endregion   

        #region SetWorkbookVisibility
        /// <summary>
        /// When a Excel document is embedded in for example a Word document the Workbook
        /// is set to hidden. Don't know why Microsoft does this but they do. To solve this
        /// problem we seek the WINDOW1 record in the BOF record of the stream. In there a
        /// grbit structure is located. The first bit in this structure controls the visibility
        /// of the workbook, so we check if this bit is set to 1 (hidden) and if so set it to 0.
        /// Normally a Workbook stream only contains one WINDOW record but when it is embedded
        /// it will contain 2 or more records.
        /// </summary>
        /// <param name="rootStorage">The <see cref="CFStorage">Root storage</see> of a <see cref="CompoundFile"/></param>
        /// <exception cref="OEFileIsCorrupt">Raised when the <paramref name="rootStorage"/> does not have a Workbook stream</exception>
        internal void SetWorkbookVisibility(CFStorage rootStorage)
        {
            if (!rootStorage.TryGetStream("WorkBook", out var stream))
                throw new OEFileIsCorrupt("Could not check workbook visibility because the WorkBook stream is not present");

            Logger.WriteToLog("Setting hidden Excel workbook to visible");

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
