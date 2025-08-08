﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using OfficeExtractor.Ole;
using OpenMcdf;
using SharpCompress.Archives;
using Version = OpenMcdf.Version;

// ReSharper disable VariableLengthStringHexEscapeSequence

//
// Extraction.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2025 Magic-Sessions. (www.magic-sessions.com)
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

namespace OfficeExtractor.Helpers;

/// <summary>
///     This class contain helpers method for extraction
/// </summary>
internal class Extraction
{
    #region Consts
    /// <summary>
    ///     Default name for embedded object without a name
    /// </summary>
    public const string DefaultEmbeddedObjectName = "Embedded object";
    #endregion

    #region Fields
    /// <summary>
    ///     <see cref="Excel" />
    /// </summary>
    private Excel _excel;
    #endregion

    #region Properties
    /// <summary>
    ///     Returns a reference to the Excel class when it already exists or creates a new one
    ///     when it doesn't
    /// </summary>
    private Excel Excel
    {
        get
        {
            if (_excel != null)
                return _excel;

            _excel = new Excel();
            return _excel;
        }
    }
    #endregion

    #region IsCompoundFile
    /// <summary>
    ///     Returns true is the byte array starts with a compound file identifier
    /// </summary>
    /// <param name="bytes"></param>
    /// <returns></returns>
    public bool IsCompoundFile(byte[] bytes)
    {
        if (bytes == null || bytes.Length < 2)
            return false;

        return bytes[0] == 0xD0 && bytes[1] == 0xCF;
    }
    #endregion

    #region GetFileNameFromObjectReplacementFile
    /// <summary>
    ///     Tries to extract the original filename for the OLE object out of the ObjectReplacement file
    /// </summary>
    /// <param name="zipEntry"></param>
    /// <returns></returns>
    internal string GetFileNameFromObjectReplacementFile(IArchiveEntry zipEntry)
    {
        Logger.WriteToLog("Trying to get original filename from ObjectReplacement file");

        try
        {
            using var zipEntryStream = zipEntry.OpenEntryStream();
            using var zipEntryMemoryStream = new MemoryStream();
            zipEntryStream.CopyTo(zipEntryMemoryStream);
            zipEntryMemoryStream.Position = 0x4470;
            using var binaryReader = new BinaryReader(zipEntryMemoryStream);
            while (binaryReader.BaseStream.Position != binaryReader.BaseStream.Length)
            {
                var value = binaryReader.ReadUInt16();

                // We have found the start position from where we are going to read
                // the original filename
                if (value != 0x8000 || binaryReader.PeekChar() != 0x46) continue;
                // Skip the peeked char
                zipEntryMemoryStream.Position += 2;

                // Read until we find the next 0x46 value
                while (binaryReader.BaseStream.Position != binaryReader.BaseStream.Length)
                {
                    value = binaryReader.ReadUInt16();
                    if (value != 0x46) continue;
                    // Skip the next 6 bytes
                    binaryReader.ReadBytes(6);

                    // Get the length of name string
                    var length = binaryReader.ReadUInt16();

                    // Skip the next 2 bytes
                    zipEntryMemoryStream.Position += 2;

                    // Read the filename bytes
                    var fileNameBytes = binaryReader.ReadBytes(length);
                    var fileName = Encoding.Unicode.GetString(fileNameBytes);
                    fileName = fileName.Replace("\0", string.Empty);
                    Logger.WriteToLog($"Filename '{fileName}' found");
                    return fileName;
                }
            }
        }
        catch
        {
            return null;
        }

        return null;
    }
    #endregion

    #region SaveStorageTreeToCompoundFile
    /// <summary>
    ///     This will save the complete tree from the given <paramref name="storage" /> to a new <see cref="Storage" />
    /// </summary>
    /// <param name="storage"></param>
    /// <param name="fileName">The filename with path for the new compound file</param>
    internal string SaveStorageTreeToCompoundFile(Storage storage, string fileName)
    {
        Logger.WriteToLog($"Saving storage tree to compound file '{fileName}'");

        fileName = FileManager.FileExistsMakeNew(fileName);

        using var compoundFile = RootStorage.Create(fileName, Version.V3, StorageModeFlags.Transacted);
        storage.CopyTo(compoundFile);
        compoundFile.Commit();
        compoundFile.Flush(true);

        return fileName;
    }
    #endregion

    #region SaveByteArrayToFile
    /// <summary>
    ///     Saves the <paramref name="data" /> bytes to the <paramref name="outputFile" />
    /// </summary>
    /// <param name="data">The stream as byte array</param>
    /// <param name="outputFile">The output filename with path</param>
    /// <returns></returns>
    /// <exception cref="OfficeExtractor.Exceptions.OEFileIsCorrupt">Raised when the file is corrupt</exception>
    internal string SaveByteArrayToFile(byte[] data, string outputFile)
    {
        if (string.IsNullOrWhiteSpace(outputFile))
            throw new ArgumentNullException(nameof(outputFile), "The output file name is empty or null");

        // seal-mb
        // Check name of file
        // maybe the name contains question marks, because the OLE Bin file
        // use it fore UNICODE signs
        if (Extractor.ReplaceDisallowedCharsInFileNames)
        {
            var fileName = FileManager.RemoveInvalidFileNameChars(Path.GetFileName(outputFile));
            outputFile = Path.Combine(Path.GetDirectoryName(outputFile)!, fileName);
        }

        // Because the data is stored in a stream we have no name for it so we
        // have to check the magic bytes to see with what kind of file we are dealing
        Logger.WriteToLog($"Saving byte array with length '{data.Length}' to file '{outputFile}'");

        var extension = Path.GetExtension(outputFile);

        if (string.IsNullOrEmpty(extension))
        {
            var fileType = FileTypeSelector.GetFileTypeFileInfo(data);
            if (fileType != null && !string.IsNullOrEmpty(fileType.Extension))
                outputFile += "." + fileType.Extension;

            if (fileType != null)
                extension = "." + fileType.Extension;
        }

        // Check if the output file already exists and if so make a new one
        outputFile = FileManager.FileExistsMakeNew(outputFile);

        switch (extension.ToUpperInvariant())
        {
            case ".XLS":
            case ".XLT":
            case ".XLW":
            {
                using var memoryStream = new MemoryStream(data);
                using var compoundFile = RootStorage.Open(memoryStream, StorageModeFlags.Transacted);
                Excel.SetWorkbookVisibility(compoundFile);
                compoundFile.Commit();
                compoundFile.Flush(true);
                break;
            }

            case ".XLSB":
            case ".XLSM":
            case ".XLSX":
            case ".XLTM":
            case ".XLTX":
                using (var memoryStream = new MemoryStream(data))
                {
                    var file = Excel.SetWorkbookVisibility(memoryStream);
                    File.WriteAllBytes(outputFile, file.ToArray());
                }

                break;

            default:
                File.WriteAllBytes(outputFile, data);
                break;
        }

        return outputFile;
    }
    #endregion

    #region Storage Node Content Parsing
    private string GetDelimitedStringFromData(string delimiter, ICollection<byte> data)
    {
        string delimitedString = null;
        if (!string.IsNullOrWhiteSpace(delimiter) && data is { Count: > 0 })
            // Check if data has at least the length of opening plus closing delimiter
            if (data.Count >= delimiter.Length * 2)
                // Check if data contains the delimiter
                if (delimiter.Equals(Encoding.UTF8.GetString(data.Take(delimiter.Length).ToArray())))
                    // Read the data after opening delimiter until first sign of the closing delimiter
                    delimitedString = Encoding.UTF8.GetString(data
                        .Skip(delimiter.Length)
                        .TakeWhile(b => Convert.ToChar(b) != delimiter.First())
                        .ToArray());

        return delimitedString;
    }
    #endregion

    #region SaveFromStorageNode
    /// <summary>
    ///     This method will extract and save the data from the given <see cref="RootStorage" /> node to the
    ///     <paramref name="outputFolder" />
    /// </summary>
    /// <param name="bytes">The <see cref="RootStorage" /> as a byte array</param>
    /// <param name="outputFolder">The outputFolder</param>
    /// <returns></returns>
    /// <exception cref="Exceptions.OEFileIsPasswordProtected">
    ///     Raised when a WordDocument, WorkBook or PowerPoint Document
    ///     stream is password protected
    /// </exception>
    internal string SaveFromStorageNode(byte[] bytes, string outputFolder)
    {
        using var memoryStream = new MemoryStream(bytes);
        using var compoundFile = RootStorage.Open(memoryStream);
        return SaveFromStorageNode(compoundFile, outputFolder, null);
    }

    /// <summary>
    ///     This method will extract and save the data from the given <see cref="RootStorage" /> node to the
    ///     <paramref name="outputFolder" />
    /// </summary>
    /// <param name="bytes">The <see cref="RootStorage" /> as a byte array</param>
    /// <param name="outputFolder">The outputFolder</param>
    /// <param name="fileName">The fileName to use, null when the fileName is unknown</param>
    /// <returns></returns>
    /// <exception cref="Exceptions.OEFileIsPasswordProtected">
    ///     Raised when a WordDocument, WorkBook or PowerPoint Document
    ///     stream is password protected
    /// </exception>
    internal string SaveFromStorageNode(byte[] bytes, string outputFolder, string fileName)
    {
        using var memoryStream = new MemoryStream(bytes);
        using var compoundFile = RootStorage.Open(memoryStream);
        return SaveFromStorageNode(compoundFile, outputFolder, fileName);
    }

    /// <summary>
    ///     This method will extract and save the data from the given <paramref name="storage" /> node to the
    ///     <paramref name="outputFolder" />
    /// </summary>
    /// <param name="storage">The <see cref="Storage" /> node</param>
    /// <param name="outputFolder">The outputFolder</param>
    /// <returns></returns>
    /// <exception cref="Exceptions.OEFileIsPasswordProtected">
    ///     Raised when a WordDocument, WorkBook or PowerPoint Document
    ///     stream is password protected
    /// </exception>
    internal string SaveFromStorageNode(Storage storage, string outputFolder)
    {
        return SaveFromStorageNode(storage, outputFolder, null);
    }

    /// <summary>
    ///     This method will extract and save the data from the given <paramref name="storage" /> node to the
    ///     <paramref name="outputFolder" />
    /// </summary>
    /// <param name="storage">The <see cref="Storage" /> node</param>
    /// <param name="outputFolder">The outputFolder</param>
    /// <param name="fileName">The fileName to use, null when the fileName is unknown</param>
    /// <returns>
    ///     Returns the name of the created file that or null if there was nothing to export within the given
    ///     <paramref name="storage" /> node.
    /// </returns>
    /// <exception cref="Exceptions.OEFileIsPasswordProtected">
    ///     Raised when a WordDocument, WorkBook or PowerPoint Document
    ///     stream is password protected
    /// </exception>
    public string SaveFromStorageNode(Storage storage, string outputFolder, string fileName)
    {
        Logger.WriteToLog($"Saving CFStorage to output folder '{outputFolder}' with file name {fileName}");

        if (storage.TryOpenStream("CONTENTS", out var contents))
        {
            Logger.WriteToLog("CONTENTS stream found");

            if (contents.Length <= 0)
            {
                Logger.WriteToLog("CONTENTS stream is empty");
                return null;
            }

            if (string.IsNullOrWhiteSpace(fileName)) fileName = DefaultEmbeddedObjectName;

            const string delimiter = "%DocumentOle:";

            var contentsData = contents.GetData();

            var documentOleFileName = GetDelimitedStringFromData(delimiter, contentsData);
            if (documentOleFileName == null)
                return SaveByteArrayToFile(contents.GetData(), FileManager.FileExistsMakeNew(Path.Combine(outputFolder, fileName)));
            
            if (!documentOleFileName.Equals(string.Empty))
                fileName = Path.GetFileName(documentOleFileName);

            return SaveByteArrayToFile(contentsData.Skip(delimiter.Length * 2 + documentOleFileName.Length).ToArray(), FileManager.FileExistsMakeNew(Path.Combine(outputFolder, fileName)));
        }

        if (storage.TryOpenStream("Package", out var package))
        {
            Logger.WriteToLog("Package stream found");

            if (package.Length <= 0)
            {
                Logger.WriteToLog("Package stream is empty");
                return null;
            }

            if (string.IsNullOrWhiteSpace(fileName)) fileName = DefaultEmbeddedObjectName;
            return SaveByteArrayToFile(package.GetData(), FileManager.FileExistsMakeNew(Path.Combine(outputFolder, fileName)));
        }

        if (storage.TryOpenStream("EmbeddedOdf", out var embeddedOdf))
        {
            Logger.WriteToLog("EmbeddedOdf stream found");

            // The embedded object is an Embedded ODF file
            if (embeddedOdf.Length <= 0)
            {
                Logger.WriteToLog("EmbeddedOdf stream is empty");
                return null;
            }

            if (string.IsNullOrWhiteSpace(fileName)) fileName = DefaultEmbeddedObjectName;
            return SaveByteArrayToFile(embeddedOdf.GetData(), FileManager.FileExistsMakeNew(Path.Combine(outputFolder, fileName)));
        }

        if (storage.TryOpenStream("\x0001Ole10Native", out _))
        {
            Logger.WriteToLog("Ole10Native stream found");

            var ole10Native = new Ole10Native(storage);
            Logger.WriteToLog($"Ole10Native stream format is '{ole10Native.Format}'");

            if (ole10Native.Format == OleFormat.File)
                return SaveByteArrayToFile(ole10Native.NativeData, FileManager.FileExistsMakeNew(Path.Combine(outputFolder, ole10Native.FileName)));

            Logger.WriteToLog("Ole10Native is ignored");
            return null;
        }

        if (storage.TryOpenStream("WordDocument", out _))
        {
            Logger.WriteToLog("WordDocument stream found");

            // The embedded object is a Word file
            if (string.IsNullOrWhiteSpace(fileName)) fileName = "Embedded Word document.doc";
            return SaveStorageTreeToCompoundFile(storage, FileManager.FileExistsMakeNew(Path.Combine(outputFolder, fileName)));
        }

        if (storage.TryOpenStream("Workbook", out _))
        {
            Logger.WriteToLog("Workbook stream found");

            // The embedded object is an Excel file   
            if (string.IsNullOrWhiteSpace(fileName)) fileName = "Embedded Excel document.xls";
            Excel.SetWorkbookVisibility(storage);
            return SaveStorageTreeToCompoundFile(storage, FileManager.FileExistsMakeNew(Path.Combine(outputFolder, fileName)));
        }

        if (storage.TryOpenStream("PowerPoint Document", out _))
        {
            Logger.WriteToLog("PowerPoint Document stream found");

            // The embedded object is a PowerPoint file
            if (string.IsNullOrWhiteSpace(fileName)) fileName = "Embedded PowerPoint document.ppt";
            return SaveStorageTreeToCompoundFile(storage, FileManager.FileExistsMakeNew(Path.Combine(outputFolder, fileName)));
        }

        return null;
    }
    #endregion
}