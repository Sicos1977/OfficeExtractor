using System;
using System.Collections.Generic;
using System.IO;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage;
using DocumentServices.Modules.Extractors.OfficeExtractor.Exceptions;
using DocumentServices.Modules.Extractors.OfficeExtractor.Helpers;

namespace DocumentServices.Modules.Extractors.OfficeExtractor
{
    /// <summary>
    /// This class is used to extract embedded files from Word, Excel and PowerPoint files
    /// </summary>
    public class Extractor
    {
        #region CheckFileNameAndOutputFolder
        /// <summary>
        /// Checks if the <see cref="inputFile"/> and <see cref="outputFolder"/> is valid
        /// </summary>
        /// <param name="inputFile"></param>
        /// <param name="outputFolder"></param>
        /// <exception cref="ArgumentNullException">Raised when the <see cref="inputFile"/> or <see cref="outputFolder"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <see cref="inputFile"/> does not exists</exception>
        /// <exception cref="DirectoryNotFoundException">Raised when the <see cref="outputFolder"/> does not exists</exception>
        private void CheckFileNameAndOutputFolder(string inputFile, string outputFolder)
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

        #region ExtractToFolder
        /// <summary>
        /// Extracts all the embedded object from the Microsoft Office <see cref="inputFile"/> to the 
        /// <see cref="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="inputFile">The Microsoft Office file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        /// <exception cref="ArgumentNullException">Raised when the <see cref="inputFile"/> or <see cref="outputFolder"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <see cref="inputFile"/> does not exists</exception>
        /// <exception cref="DirectoryNotFoundException">Raised when the <see cref="outputFolder"/> does not exists</exception>
        /// <exception cref="OEFileTypeNotSupported">Raised when the Microsoft Office File Type is not supported</exception>
        public List<string> ExtractToFolder(string inputFile, string outputFolder)
        {
            CheckFileNameAndOutputFolder(inputFile, outputFolder);
            
            var extension = Path.GetExtension(inputFile);
            if (extension != null)
                extension = extension.ToUpperInvariant();

            outputFolder = FileManager.CheckForBackSlash(outputFolder);

            switch (extension)
            {
                case ".DOC":
                case ".DOT":
                    // Word 97 - 2003
                    return ExtractFromWordBinaryFormat(inputFile, outputFolder);

                case ".DOCM":
                case ".DOCX":
                case ".DOTM":
                    // Word 2007 - 2013
                    return ExtractFromOfficeOpenXmlFormat(inputFile, "/word/embeddings/", outputFolder);

                case ".XLS":
                case ".XLT":
                case ".XLW":
                    // Excel 97 - 2003
                    return ExtractFromExcelBinaryFormat(inputFile, outputFolder);

                case ".XLSB":
                case ".XLSM":
                case ".XLSX":
                case ".XLTM":
                case ".XLTX":
                    // Excel 2007 - 2013
                    return ExtractFromOfficeOpenXmlFormat(inputFile, "/excel/embeddings/", outputFolder);

                case ".POT":
                case ".PPS":
                    // PowerPoint 97 - 2003
                    return ExtractFromPowerPointBinaryFormat(inputFile, outputFolder);

                case ".POTM":
                case ".POTX":
                case ".PPSM":
                case ".PPSX":
                case ".PPTM":
                case ".PPTX":
                    // PowerPoint 2007 - 2013
                    return ExtractFromOfficeOpenXmlFormat(inputFile, "/ppt/embeddings/", outputFolder);

                default:
                    throw new OEFileTypeNotSupported("The file '" + Path.GetFileName(inputFile) +
                                                     "' is not supported, only .DOC, .DOCM, .DOCX, .DOT, .DOTM, .XLS, .XLSB, .XLSM, .XLSX, .XLT, .XLTM, .XLTX, .XLW, .POT, .POTM, .POTX, .PPS, .PPSM, .PPSX, .PPTM and .PPTX are supported");
            }
        }
        #endregion

        #region ExtractFromWordBinaryFormat
        /// <summary>
        /// Extracts all the embedded Word object from the <see cref="inputFile"/> to the 
        /// <see cref="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="inputFile">The binary Word file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        public List<string> ExtractFromWordBinaryFormat(string inputFile, string outputFolder)
        {
            var result = new List<string>();

            var compoundFile = new CompoundFile(inputFile);

            // In a Word file the objects are stored in the ObjectPool tree
            var objectPools = compoundFile.GetAllNamedEntries("ObjectPool");
            foreach (var objectPool in objectPools)
            {
                // An objectPool is always a CFStorage type
                var objectPoolStorage = objectPool as CFStorage;
                if (objectPoolStorage == null) continue;

                // Multiple objects are stored as children of the objectPool
                foreach (var child in objectPoolStorage.Children)
                {
                    var childStorage = child as CFStorage;
                    if (childStorage == null) continue;

                    // Embedded objects can be stored in 4 ways
                    // - As a CONTENT stream
                    // - As a Package
                    // - As an Ole10Native object
                    // - Embedded into the same compound file
                    if (childStorage.ExistsStream("CONTENTS"))
                    {
                        var contents = childStorage.GetStream("CONTENTS");
                        if (contents.Size > 0)
                            result.Add(SaveByteArrayToFile(contents.GetData(), outputFolder + "embedded word object"));
                    }
                    else if (childStorage.ExistsStream("Package"))
                    {
                        var package = childStorage.GetStream("Package");
                        if (package.Size > 0)
                            result.Add(SaveByteArrayToFile(package.GetData(), outputFolder + "embedded word object"));
                    }
                    else if (childStorage.ExistsStream("\x01Ole10Native"))
                    {
                        var ole10Native = childStorage.GetStream("\x01Ole10Native");
                        if (ole10Native.Size > 0)
                            result.Add(ExtractFileFromOle10Native(ole10Native.GetData(), outputFolder));
                    }
                    else if (childStorage.ExistsStream("WordDocument"))
                    {
                        // The embedded object is a Word file
                        var tempFileName = outputFolder + "Embedded Word document.doc";
                        compoundFile.SaveNamedEntryTreeToCompoundFile(childStorage, tempFileName);
                        result.Add(tempFileName);
                    }
                    else if (childStorage.ExistsStream("Workbook"))
                    {
                        // The embedded object is an Excel file   
                        var tempFileName = outputFolder + "Embedded Excel document.xls";
                        compoundFile.SaveNamedEntryTreeToCompoundFile(childStorage, tempFileName);
                        result.Add(tempFileName);
                    }
                    else if (childStorage.ExistsStream("PowerPoint Document"))
                    {
                        // The embedded object is a PowerPoint file
                        var tempFileName = outputFolder + "Embedded PowerPoint document.ppt";
                        compoundFile.SaveNamedEntryTreeToCompoundFile(childStorage, tempFileName);
                        result.Add(tempFileName);
                    }
                }
            }

            return result;
        }
        #endregion

        #region ExtractFromExcelBinaryFormat
        /// <summary>
        /// Extracts all the embedded Excel object from the <see cref="inputFile"/> to the 
        /// <see cref="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="inputFile">The binary Excel file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        public List<string> ExtractFromExcelBinaryFormat(string inputFile, string outputFolder)
        {
            throw new NotImplementedException("Not yet fully implemented");

            var result = new List<string>();

            var compoundFile = new CompoundFile(inputFile);
            // In a Word file the objects are stored in the ObjectPool tree
            var objectPools = compoundFile.GetAllNamedEntries("ObjectPool");
            foreach (var objectPool in objectPools)
            {
                // An objectPool is always a CFStorage type
                var objectPoolStorage = objectPool as CFStorage;
                if (objectPoolStorage == null) continue;

                // Multiple objects are stored as children of the objectPool
                foreach (var child in objectPoolStorage.Children)
                {
                    var childStorage = child as CFStorage;
                    if (childStorage == null) continue;

                    // Ole objects can be stored in 4 ways
                    // - As a CONTENT stream
                    // - As a Package
                    // - As an Ole10Native object
                    // - Embedded into the same compound file
                    if (childStorage.ExistsStream("CONTENTS"))
                    {
                        var contents = childStorage.GetStream("CONTENTS");
                        if (contents.Size > 0)
                            result.Add(SaveByteArrayToFile(contents.GetData(), outputFolder + "embedded word object"));
                    }
                    else if (childStorage.ExistsStream("Package"))
                    {
                        var package = childStorage.GetStream("Package");
                        if (package.Size > 0)
                            result.Add(SaveByteArrayToFile(package.GetData(), outputFolder + "embedded word object"));
                    }
                    else if (childStorage.ExistsStream("\x01Ole10Native"))
                    {
                        var ole10Native = childStorage.GetStream("\x01Ole10Native");
                        if (ole10Native.Size > 0)
                            result.Add(ExtractFileFromOle10Native(ole10Native.GetData(), outputFolder));
                    }

                    // Workbook
                    // PowerPoint Document
                    // WordDocument
                }
            }

            return result;
        }
        #endregion

        #region ExtractFromPowerPointBinaryFormat
        /// <summary>
        /// Extracts all the embedded PowerPoint object from the <see cref="inputFile"/> to the 
        /// <see cref="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="inputFile">The binary Excel file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        public List<string> ExtractFromPowerPointBinaryFormat(string inputFile, string outputFolder)
        {
            throw new NotImplementedException("Not yet fully implemented");

            var result = new List<string>();

            var compoundFile = new CompoundFile(inputFile);
            // In a Word file the objects are stored in the ObjectPool tree
            var objectPools = compoundFile.GetAllNamedEntries("ObjectPool");
            foreach (var objectPool in objectPools)
            {
                // An objectPool is always a CFStorage type
                var objectPoolStorage = objectPool as CFStorage;
                if (objectPoolStorage == null) continue;

                // Multiple objects are stored as children of the objectPool
                foreach (var child in objectPoolStorage.Children)
                {
                    var childStorage = child as CFStorage;
                    if (childStorage == null) continue;

                    // Ole objects can be stored in 4 ways
                    // - As a CONTENT stream
                    // - As a Package
                    // - As an Ole10Native object
                    // - Embedded into the same compound file
                    if (childStorage.ExistsStream("CONTENTS"))
                    {
                        var contents = childStorage.GetStream("CONTENTS");
                        if (contents.Size > 0)
                            result.Add(SaveByteArrayToFile(contents.GetData(), outputFolder + "embedded word object"));
                    }
                    else if (childStorage.ExistsStream("Package"))
                    {
                        var package = childStorage.GetStream("Package");
                        if (package.Size > 0)
                            result.Add(SaveByteArrayToFile(package.GetData(), outputFolder + "embedded word object"));
                    }
                    else if (childStorage.ExistsStream("\x01Ole10Native"))
                    {
                        var ole10Native = childStorage.GetStream("\x01Ole10Native");
                        if (ole10Native.Size > 0)
                            result.Add(ExtractFileFromOle10Native(ole10Native.GetData(), outputFolder));
                    }

                    // Workbook
                    // PowerPoint Document
                    // WordDocument
                }
            }

            return result;
        }
        #endregion

        #region ExtractFromOfficeOpenXmlFormat
        /// <summary>
        /// Extracts all the embedded object from the Office Open XML <see cref="inputFile"/> to the 
        /// <see cref="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="inputFile">The Office Open XML format file</param>
        /// <param name="zipFolder">The folder in the Office Open XML format (zip) file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        public List<string> ExtractFromOfficeOpenXmlFormat(string inputFile, string zipFolder, string outputFolder)
        {
            throw new NotImplementedException("Not yet implemented");
            //Package pkg = Package.Open(fileName);


            //// Get the embedded files names. 
            //foreach (PackagePart pkgPart in pkg.GetParts())
            //{
            //    if (pkgPart.Uri.ToString().StartsWith(embeddingPartString))
            //    {
            //        string fileName1 = pkgPart.Uri.ToString().Remove(0, embeddingPartString.Length);
            //        chkdLstEmbeddedFiles.Items.Add(fileName1);
            //    }
            //}
            //pkg.Close(); 

        }
        #endregion

        #region SaveByteArrayToFile
        /// <summary>
        /// Saves the <see cref="data"/> byte array to the <see cref="outputFile"/>
        /// </summary>
        /// <param name="data">The stream as byte array</param>
        /// <param name="outputFile">The output filename with path</param>
        /// <returns></returns>
        private static string SaveByteArrayToFile(byte[] data, string outputFile)
        {
            // Because the data is stored in a stream we have no name for it so we
            // have to check the magic bytes to see with what kind of file we are dealing
            var fileType = FileTypeSelector.GetFileTypeFileInfo(data);
            if (fileType != null && !string.IsNullOrEmpty(fileType.Extension))
                outputFile += "." + fileType.Extension;

            // Check if the output file already exists and if so make a new one
            outputFile = FileManager.FileExistsMakeNew(outputFile);

            File.WriteAllBytes(outputFile, data);
            return outputFile;
        }
        #endregion

        #region ExtractFileFromOle10Native
        /// <summary>
        /// Extract the file from the Ole10Native container and saves it to the outputfolder
        /// </summary>
        /// <param name="ole10Native">The Ole10Native object as an byte array</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>The filename with path from the extracted file</returns>
        private static string ExtractFileFromOle10Native(byte[] ole10Native, string outputFolder)
        {
            // Convert the byte array to a stream
            using (Stream oleStream = new MemoryStream(ole10Native))
            {
                // The name of the file start at postion 7 so move to there
                oleStream.Position = 6;
                var tempFileName = new char[260];

                // Read until we find a null character
                int i;
                var chr = new byte[1];
                for (i = 0; i < 260; i++)
                {
                    oleStream.Read(chr, 0, 1);
                    tempFileName[i] = (char) chr[0];
                    if (chr[0] == 0)
                        break;
                }

                var fileName = new string(tempFileName, 0, i);

                // We don't need this but we need to read it to know where we
                // are located in the stream
                var tempOriginalFilePath = new char[260];
                for (i = 0; i < 260; i++)
                {
                    oleStream.Read(chr, 0, 1);
                    tempOriginalFilePath[i] = (char) chr[0];
                    if (chr[0] == 0)
                        break;
                }

                //var originalFilePath = new string(tempOriginalFilePath, 0, i);

                // We need to skip the next four bytes
                oleStream.Position += 4;

                // Read the tempory path size
                var size = new byte[4];
                oleStream.Read(size, 0, 4);
                var tempPathSize = BitConverter.ToInt32(size, 0);

                // Move the position in the stream after the temp path
                oleStream.Position += tempPathSize;

                // Read the next four bytes for the length of the data
                oleStream.Read(size, 0, 4);
                var fileSize = BitConverter.ToInt32(size, 0);

                // And finaly we have come to the original file
                var fileData = new byte[fileSize];
                oleStream.Read(fileData, 0, fileSize);

                // Check if the output file already exists and if so make a new one
                fileName = outputFolder + fileName;
                fileName = FileManager.FileExistsMakeNew(fileName);

                File.WriteAllBytes(fileName, fileData);
                return fileName;
            }
        }
        #endregion
    }
}
