using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
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
// Copyright (c) 2013-2026 Kees van Spelde. (www.magic-sessions.com)
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

namespace OfficeExtractor;

/// <summary>
///     This class is used as a placeholder for all Excel related methods
/// </summary>
internal class Excel : OfficeBase
{
    #region Extract
    /// <summary>
    ///     This method saves all the Excel embedded binary objects from the <paramref name="inputFile" /> to the
    ///     <paramref name="outputFolder" />
    /// </summary>
    /// <param name="inputFile">The binary Excel file</param>
    /// <param name="outputFolder">The output folder</param>
    /// <returns></returns>
    /// <exception cref="OEFileIsPasswordProtected">Raised when the <paramref name="inputFile" /> is password protected</exception>
    /// <exception cref="OEFileIsCorrupt">Raised when the file is corrupt</exception>
    internal List<string> Extract(string inputFile, string outputFolder, bool continueOnError = false)
    {
        Logger.WriteToLog("The file is a binary Excel sheet");

        using var compoundFile = RootStorage.OpenRead(inputFile);
        var result = new List<string>();

        foreach (var item in compoundFile.EnumerateEntries())
        {
            try
            {
                if (!item.Name.StartsWith("MBD")) continue;
                if (!compoundFile.TryOpenStorage(item.Name, out var storage)) continue;
                var extractedFileName = Extraction.SaveFromStorageNode(storage, outputFolder);
                if (extractedFileName != null) result.Add(extractedFileName);
            }
            catch (Exception ex)
            {
                HandleException(ex, "Excel", shallThrow: !continueOnError);
            }
        }

        return result;
    }
    #endregion

    #region SetWorkbookVisibility
    /// <summary>
    ///     When an Excel document is embedded in for example a Word document the Workbook
    ///     is set to hidden. Don't know why Microsoft does this, but they do. To solve this
    ///     problem we seek the WINDOW1 record in the BOF record of the stream. In there a
    ///     grbit structure is located. The first bit in this structure controls the visibility
    ///     of the workbook, so we check if this bit is set to 1 (hidden) and if so set it to 0.
    ///     Normally a Workbook stream only contains one WINDOW record but when it is embedded
    ///     it will contain 2 or more records.
    /// </summary>
    /// <param name="rootStorage">The <see cref="Storage">Root storage</see> of a <see cref="RootStorage" /></param>
    /// <exception cref="OEFileIsCorrupt">Raised when the <paramref name="rootStorage" /> does not have a Workbook stream</exception>
    internal void SetWorkbookVisibility(Storage rootStorage)
    {
        if (!rootStorage.TryOpenStream("WorkBook", out var stream))
            throw new OEFileIsCorrupt("Could not check workbook visibility because the WorkBook stream is not present");

        Logger.WriteToLog("Setting hidden Excel workbook to visible");

        try
        {
            using (var binaryReader = new BinaryReader(stream, Encoding.Default, leaveOpen: true))
            using (var binaryWriter = new BinaryWriter(stream, Encoding.Default, leaveOpen: true))
            {
                var recordType = binaryReader.ReadUInt16();
                while (binaryReader.BaseStream.Position < binaryReader.BaseStream.Length)
                {
                    var recordStartPos = binaryReader.BaseStream.Position;

                    var recordLength = binaryReader.ReadUInt16();

                    var grbit = binaryReader.ReadBytes(2);
                    if (recordType == 0x3D) // Window1 record
                    {
                        // Skip xWn, yWn, dxWn, dyWn (4 × 2 bytes)
                        binaryReader.ReadUInt16(); // xWn
                        binaryReader.ReadUInt16(); // yWn
                        binaryReader.ReadUInt16(); // dxWn
                        binaryReader.ReadUInt16(); // dyWn

                        var grbitPos = binaryReader.BaseStream.Position;

                        var bitArray = new BitArray(grbit);

                        if (bitArray.Get(0))
                        {
                            bitArray.Set(0, false);

                            var modified = new byte[2];
                            bitArray.CopyTo(modified, 0);
                            binaryWriter.BaseStream.Position = grbitPos;
                            binaryWriter.Write(modified);
                        }

                        break;
                    }

                    binaryReader.BaseStream.Position = recordStartPos + 4 + recordLength; // Skip this record
                }
            }

            stream.Position = 0;
        }
        catch (Exception exception)
        {
            throw new OEFileIsCorrupt("Could not check workbook visibility because the file seems to be corrupt", exception);
        }
    }

    /// <summary>
    ///     This method sets the workbook in an Open XML Format Excel file to visible
    /// </summary>
    /// <param name="spreadSheetDocument">The Open XML Format Excel file as a memory stream</param>
    /// <exception cref="OEFileIsCorrupt">Raised when the <paramref name="spreadSheetDocument" /> is corrupt</exception>
    public static MemoryStream SetWorkbookVisibility(MemoryStream spreadSheetDocument)
    {
        try
        {
            using var spreadsheetDocument = SpreadsheetDocument.Open(spreadSheetDocument, true);
            // ReSharper disable PossibleNullReferenceException
            var bookViews = spreadsheetDocument.WorkbookPart.Workbook.BookViews;
            foreach (var bookView in bookViews)
            {
                var workBookView = (WorkbookView)bookView;
                if (workBookView.Visibility.Value == VisibilityValues.Hidden ||
                    workBookView.Visibility.Value == VisibilityValues.VeryHidden)
                    workBookView.Visibility.Value = VisibilityValues.Visible;
            }

            spreadsheetDocument.WorkbookPart.Workbook.Save();
            // ReSharper restore PossibleNullReferenceException

            return spreadSheetDocument;
        }
        catch (Exception exception)
        {
            throw new OEFileIsCorrupt("Could not check workbook visibility because the file seems to be corrupt",
                exception);
        }
    }
    #endregion
}