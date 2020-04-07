using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;
using OpenMcdf;
using PasswordProtectedChecker;

//
// Extractor.cs
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
    /// This class is used to extract embedded files from Word, Excel and PowerPoint files. It only extracts
    /// one level deep, for example when you have an Word file with an embedded Excel file that has an embedded
    /// PDF it will only extract the embedded Excel file from the Word file.
    /// </summary>
    public class Extractor
    {
        #region Fields
        /// <summary>
        ///     <see cref="Checker"/>
        /// </summary>
        private readonly Checker _passwordProtectedChecker = new Checker();

        /// <summary>
        ///     <see cref="Word"/>
        /// </summary>
        private Word _word;

        /// <summary>
        ///     <see cref="Excel"/>
        /// </summary>
        private Excel _excel;

        /// <summary>
        ///     <see cref="PowerPoint"/>
        /// </summary>
        private PowerPoint _powerPoint;

        /// <summary>
        ///     <see cref="Rtf"/>
        /// </summary>
        private Rtf _rtf;

        /// <summary>
        ///     <see cref="Extraction"/>
        /// </summary>
        private Extraction _extraction;
        #endregion

        #region Properties
        /// <summary>
        ///     An unique id that can be used to identify the logging of the converter when
        ///     calling the code from multiple threads and writing all the logging to the same file
        /// </summary>
        // ReSharper disable once UnusedMember.Global
        public string InstanceId
        {
            get => Logger.InstanceId;
            set => Logger.InstanceId = value;
        }

        /// <summary>
        /// Returns a reference to the Word class when it already exists or creates a new one
        /// when it doesn't
        /// </summary>
        private Word Word
        {
            get
            {
                if (_word != null)
                    return _word;

                _word = new Word();
                return _word;
            }
        }

        /// <summary>
        /// Returns a reference to the Excel class when it already exists or creates a new one
        /// when it doesn't
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

        /// <summary>
        /// Returns a reference to the PowerPoint class when it already exists or creates a new one
        /// when it doesn't
        /// </summary>
        private PowerPoint PowerPoint
        {
            get
            {
                if (_powerPoint != null)
                    return _powerPoint;

                _powerPoint = new PowerPoint();
                return _powerPoint;
            }
        }

        /// <summary>
        /// Returns a reference to the RTF class when it already exists or creates a new one
        /// when it doesn't
        /// </summary>
        private Rtf Rtf
        {
            get
            {
                if (_rtf != null)
                    return _rtf;

                _rtf = new Rtf();
                return _rtf;
            }
        }

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

        #region Constructor
        /// <summary>
        ///     Creates this object and sets it's needed properties
        /// </summary>
        /// <param name="logStream">When set then logging is written to this stream for all extractions. If
        /// you want a separate log for each conversion then set the logstream on the <see cref="Extract"/> method</param>
        public Extractor(Stream logStream = null)
        {
            Logger.LogStream = logStream;
        }
        #endregion

        #region CheckFileNameAndOutputFolder
        /// <summary>
        /// Checks if the <paramref name="inputFile"/> and <paramref name="outputFolder"/> is valid
        /// </summary>
        /// <param name="inputFile"></param>
        /// <param name="outputFolder"></param>
        /// <exception cref="ArgumentNullException">Raised when the <paramref name="inputFile"/> or <paramref name="outputFolder"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <paramref name="inputFile"/> does not exists</exception>
        /// <exception cref="DirectoryNotFoundException">Raised when the <paramref name="outputFolder"/> does not exists</exception>
        private static void CheckFileNameAndOutputFolder(string inputFile, string outputFolder)
        {
            if (string.IsNullOrEmpty(inputFile))
                throw new ArgumentNullException(inputFile);

            if (string.IsNullOrEmpty(outputFolder))
                throw new ArgumentNullException(outputFolder);

            if (!File.Exists(inputFile))
                throw new FileNotFoundException(inputFile);

            if (!Directory.Exists(outputFolder))
                throw new DirectoryNotFoundException(outputFolder);
        }
        #endregion

        #region GetExtension
        /// <summary>
        ///     Get the extension from the file and checks if this extension is valid
        /// </summary>
        /// <param name="inputFile">The file to check</param>
        /// <returns></returns>
        private static string GetExtension(string inputFile)
        {
            var extension = Path.GetExtension(inputFile);
            extension = string.IsNullOrEmpty(extension) ? string.Empty : extension.ToUpperInvariant();

            switch (extension)
            {
                case ".RTF":
                case ".ODT":
                case ".ODS":
                case ".ODP":
                    break;

                default:

                    using (var fileStream = File.OpenRead(inputFile))
                    {
                        var header = new byte[2];
                        fileStream.Read(header, 0, 2);

                        // 50 4B = PK --> .doc = 4
                        if (header[0] == 0x50 && header[1] == 0x4B && extension.Length == 4)
                        {
                            extension += "X";
                        }
                        // D0 CF = DI --> .docx = 5
                        else if (header[0] == 0xD0 && header[1] == 0xCF)
                        {
                            extension = extension.Substring(0, 4);
                        }
                    }
                    break;
            }

            return extension;
        }
        #endregion

        #region ThrowPasswordProtected
        private void ThrowPasswordProtected(string inputFile)
        {
            var message = "The file '" + Path.GetFileName(inputFile) +
                          "' is password protected";
            Logger.WriteToLog(message);
            throw new OEFileIsPasswordProtected(message);
        }
        #endregion

        #region Extract
        /// <summary>
        /// Extracts all the embedded object from the Microsoft Office <paramref name="inputFile"/> to the 
        /// <paramref name="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="inputFile">The Microsoft Office file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <param name="logStream">When set then logging is written to this stream</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        /// <exception cref="ArgumentNullException">Raised when the <paramref name="inputFile"/> or <paramref name="outputFolder"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <sparamref name="inputFile"/> does not exist</exception>
        /// <exception cref="DirectoryNotFoundException">Raised when the <paramref name="outputFolder"/> does not exists</exception>
        /// <exception cref="OEFileIsCorrupt">Raised when the <paramref name="inputFile" /> is corrupt</exception>
        /// <exception cref="OEFileTypeNotSupported">Raised when the <paramref name="inputFile"/> is not supported</exception>
        /// <exception cref="OEFileIsPasswordProtected">Raised when the <paramref name="inputFile"/> is password protected</exception>
        public List<string> Extract(string inputFile, string outputFolder, Stream logStream = null)
        {
            if (logStream != null)
                Logger.LogStream = logStream;

            CheckFileNameAndOutputFolder(inputFile, outputFolder);

            var extension = GetExtension(inputFile);

            Logger.WriteToLog($"Checking if file '{inputFile}' contains any embeded objects");
            
            outputFolder = FileManager.CheckForDirectorySeparator(outputFolder);

            try
            {
                switch (extension)
                {
                    case ".ODT":
                    case ".ODS":
                    case ".ODP":
                    {
                        var result = _passwordProtectedChecker.IsFileProtected(inputFile);
                        if (result.Protected)
                            ThrowPasswordProtected(inputFile);

                        return ExtractFromOpenDocumentFormat(inputFile, outputFolder);
                    }

                    case ".DOC":
                    case ".DOT":
                    {
                        var result = _passwordProtectedChecker.IsFileProtected(inputFile);
                        if (result.Protected)
                            ThrowPasswordProtected(inputFile);

                        // Word 97 - 2003
                        return Word.Extract(inputFile, outputFolder);
                    }

                    case ".DOCM":
                    case ".DOCX":
                    case ".DOTM":
                    case ".DOTX":
                    {
                        var result = _passwordProtectedChecker.IsFileProtected(inputFile);
                        if (result.Protected)
                            ThrowPasswordProtected(inputFile);

                        // Word 2007 - 2013
                        return ExtractFromOfficeOpenXmlFormat(inputFile, "/word/embeddings/", outputFolder);
                    }

                    case ".RTF":
                        return Rtf.Extract(inputFile, outputFolder);

                    case ".XLS":
                    case ".XLT":
                    case ".XLW":
                    {
                        var result = _passwordProtectedChecker.IsFileProtected(inputFile);
                        if (result.Protected)
                            ThrowPasswordProtected(inputFile);

                        // Excel 97 - 2003
                        return Excel.Extract(inputFile, outputFolder);
                    }

                    case ".XLSB":
                    case ".XLSM":
                    case ".XLSX":
                    case ".XLTM":
                    case ".XLTX":
                    {
                        var result = _passwordProtectedChecker.IsFileProtected(inputFile);
                        if (result.Protected)
                            ThrowPasswordProtected(inputFile);

                        // Excel 2007 - 2013
                        return ExtractFromOfficeOpenXmlFormat(inputFile, "/xl/embeddings/", outputFolder);
                    }

                    case ".POT":
                    case ".PPT":
                    case ".PPS":
                    {
                        var result = _passwordProtectedChecker.IsFileProtected(inputFile);
                        if (result.Protected)
                            ThrowPasswordProtected(inputFile);

                        // PowerPoint 97 - 2003
                        return PowerPoint.Extract(inputFile, outputFolder);
                    }

                    case ".POTM":
                    case ".POTX":
                    case ".PPSM":
                    case ".PPSX":
                    case ".PPTM":
                    case ".PPTX":
                    {
                        var result = _passwordProtectedChecker.IsFileProtected(inputFile);
                        if (result.Protected)
                            ThrowPasswordProtected(inputFile);

                        // PowerPoint 2007 - 2013
                        return ExtractFromOfficeOpenXmlFormat(inputFile, "/ppt/embeddings/", outputFolder);
                    }

                    default:
                        var message = "The file '" + Path.GetFileName(inputFile) +
                                      "' is not supported, only .ODT, .DOC, .DOCM, .DOCX, .DOT, .DOTM, .DOTX, .RTF, .XLS, .XLSB, .XLSM, .XLSX, .XLT, " +
                                      ".XLTM, .XLTX, .XLW, .POT, .PPT, .POTM, .POTX, .PPS, .PPSM, .PPSX, .PPTM and .PPTX are supported";

                        Logger.WriteToLog(message);
                        throw new OEFileTypeNotSupported(message);
                }
            }
            catch (CFCorruptedFileException)
            {
                throw new OEFileIsCorrupt("The file '" + Path.GetFileName(inputFile) + "' is corrupt");
            }
            catch (Exception exception)
            {
                Logger.WriteToLog($"Cant check for embedded object because an error occured, error: {exception.Message}");
            }

            return new List<string>();
        }
        #endregion

        #region ExtractFromOfficeOpenXmlFormat
        /// <summary>
        /// Extracts all the embedded object from the Office Open XML <paramref name="inputFile"/> to the 
        /// <paramref name="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="inputFile">The Office Open XML format file</param>
        /// <param name="embeddingPartString">The folder in the Office Open XML format (zip) file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>List with files or an empty list when there are nog embedded files</returns>
        /// <exception cref="OEFileIsPasswordProtected">Raised when the Microsoft Office file is password protected</exception>
        internal List<string> ExtractFromOfficeOpenXmlFormat(string inputFile, string embeddingPartString, string outputFolder)
        {
            Logger.WriteToLog("The file is of type 'Open XML format'");

            var result = new List<string>();

            using (var inputFileMemoryStream = new MemoryStream(File.ReadAllBytes(inputFile)))
            {
                try
                {
                    var package = Package.Open(inputFileMemoryStream);

                    // Get the embedded files names. 
                    foreach (var packagePart in package.GetParts())
                    {
                        if (packagePart.Uri.ToString().StartsWith(embeddingPartString))
                        {
                            using (var packagePartStream = packagePart.GetStream())
                            using (var packagePartMemoryStream = new MemoryStream())
                            {
                                packagePartStream.CopyTo(packagePartMemoryStream);

                                var fileName = outputFolder +
                                               packagePart.Uri.ToString().Remove(0, embeddingPartString.Length);
                                
                                if (fileName.ToUpperInvariant().Contains("OLEOBJECT"))
                                {
                                    Logger.WriteToLog("OLEOBJECT found");
                                    
                                    using (var compoundFile = new CompoundFile(packagePartStream))
                                    {
                                        var resultFileName = Extraction.SaveFromStorageNode(compoundFile.RootStorage, outputFolder);
                                        if (resultFileName != null)
                                            result.Add(resultFileName);
                                        //result.Add(ExtractFileFromOle10Native(packagePartMemoryStream.ToArray(), outputFolder));
                                    }
                                }
                                else
                                {
                                    fileName = FileManager.FileExistsMakeNew(fileName);
                                    File.WriteAllBytes(fileName, packagePartMemoryStream.ToArray());
                                    result.Add(fileName);
                                }
                            }
                        }
                    }
                    package.Close();

                    return result;
                }
                catch (FileFormatException fileFormatException)
                {
                    if (!fileFormatException.Message.Equals("File contains corrupted data.",
                            StringComparison.InvariantCultureIgnoreCase))
                        return result;
                }
            }

            return result;
        }
        #endregion

        #region ExtractFromOpenDocumentFormat
        /// <summary>
        /// Searches for the first archive entry with the given name in the given archive.
        /// </summary>
        /// <param name="archive">The archive where the entry should be searched.</param>
        /// <param name="entryName">The name of the entry, which is the file or directory name.
        /// The search is done case insensitive.</param>
        /// <returns>Returns the reference of the entry if found and null if the entry doesn't exists in the archive.</returns>
        internal SharpCompress.Archives.IArchiveEntry FindEntryByName(SharpCompress.Archives.IArchive archive, string entryName)
        {
            try
            {
                return
                    archive.Entries.First(
                        archiveEntry => archiveEntry.Key.Equals(entryName, StringComparison.OrdinalIgnoreCase));
            }
            catch (InvalidOperationException)
            {
                return null;
            }
        }

        /// <summary>
        /// Extracts all the embedded object from the OpenDocument <paramref name="inputFile"/> to the 
        /// <paramref name="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="inputFile">The OpenDocument format file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        /// <exception cref="OEFileIsPasswordProtected">Raised when the OpenDocument format file is password protected</exception>
        internal List<string> ExtractFromOpenDocumentFormat(string inputFile, string outputFolder)
        {
            Logger.WriteToLog("The file is of type 'Open document format'");

            var result = new List<string>();
            using(var zipFile = SharpCompress.Archives.Zip.ZipArchive.Open(inputFile))
            {
                // Check if the file is password protected
                var manifestEntry = FindEntryByName(zipFile, "META-INF/manifest.xml");
                if (manifestEntry != null)
                {
                    using (var manifestEntryStream = manifestEntry.OpenEntryStream())
                    using (var manifestEntryMemoryStream = new MemoryStream())
                    {
                        manifestEntryStream.CopyTo(manifestEntryMemoryStream);
                        manifestEntryMemoryStream.Position = 0;
                        using (var streamReader = new StreamReader(manifestEntryMemoryStream))
                        {
                            var manifest = streamReader.ReadToEnd();
                            if (manifest.ToUpperInvariant().Contains("ENCRYPTION-DATA"))
                                throw new OEFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                                    "' is password protected");
                        }
                    }
                }

                foreach (var zipEntry in zipFile.Entries)
                {
                    if (zipEntry.IsDirectory) continue;
                    if (zipEntry.IsEncrypted)
                        throw new OEFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                            "' is password protected");

                    var name = zipEntry.Key.ToUpperInvariant();
                    if (!name.StartsWith("OBJECT") || name.Contains("/"))
                        continue;

                    string fileName = null;

                    var objectReplacementFile = FindEntryByName(zipFile, "ObjectReplacements/" + name);
                    if (objectReplacementFile != null)
                        fileName = Extraction.GetFileNameFromObjectReplacementFile(objectReplacementFile);

                    Logger.WriteToLog($"Extracting embedded object '{fileName}'");

                    using (var zipEntryStream = zipEntry.OpenEntryStream())
                    using (var zipEntryMemoryStream = new MemoryStream())
                    {
                        zipEntryStream.CopyTo(zipEntryMemoryStream);

                        using (var compoundFile = new CompoundFile(zipEntryMemoryStream))
                            result.Add(Extraction.SaveFromStorageNode(compoundFile.RootStorage, outputFolder,
                                fileName));
                    }
                }
            }
            return result;
        }
        #endregion
    }
}