using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
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
                    return ExtractFromWordBinaryFormat(inputFile, outputFolder, "ObjectPool");

                case ".DOCM":
                case ".DOCX":
                case ".DOTM":
                    // Word 2007 - 2013
                    return ExtractFromOfficeOpenXmlFormat(inputFile, "/word/embeddings/", outputFolder);

                case ".XLS":
                case ".XLT":
                case ".XLW":
                    // Excel 97 - 2003
                    return ExtractFromExcelBinaryFormat(inputFile, outputFolder, "MBD");

                case ".XLSB":
                case ".XLSM":
                case ".XLSX":
                case ".XLTM":
                case ".XLTX":
                    // Excel 2007 - 2013
                    return ExtractFromOfficeOpenXmlFormat(inputFile, "/excel/embeddings/", outputFolder);

                case ".POT":
                case ".PPT":
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
                                                     "' is not supported, only .DOC, .DOCM, .DOCX, .DOT, .DOTM, .XLS, .XLSB, .XLSM, .XLSX, .XLT, " +
                                                     ".XLTM, .XLTX, .XLW, .POT, .PPT, .POTM, .POTX, .PPS, .PPSM, .PPSX, .PPTM and .PPTX are supported");
            }
        }
        #endregion

        #region ExtractFromWordBinaryFormat
        /// <summary>
        /// This method saves all the Word embedded binary objects from the <see cref="inputFile"/> to the
        /// <see cref="outputFolder"/>
        /// </summary>
        /// <param name="inputFile">The binary Word file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <param name="storageName">The complete or part of the name from the storage that needs to be saved</param>
        /// <returns></returns>
        private List<string> ExtractFromWordBinaryFormat(string inputFile, string outputFolder, string storageName)
        {
            var compoundFile = new CompoundFile(inputFile);
            
            var result = new List<string>();

            if (compoundFile.RootStorage.ExistsStorage("ObjectPool"))
            {
                var objectPoolStorage = compoundFile.RootStorage.GetStorage("ObjectPool") as CFStorage;
                if (objectPoolStorage != null)
                {
                    // Multiple objects are stored as children of the storage object
                    foreach (var child in objectPoolStorage.Children)
                    {
                        var childStorage = child as CFStorage;
                        if (childStorage != null)
                        {
                            var extractedFileName = ExtractFromStorageNode(compoundFile, childStorage, outputFolder);
                            if (extractedFileName != null)
                                result.Add(extractedFileName);
                        }
                    }
                }
            }

            return result;
        }
        #endregion

        #region ExtractFromExcelBinaryFormat
        /// <summary>
        /// This method saves all the Excel embedded binary objects from the <see cref="inputFile"/> to the
        /// <see cref="outputFolder"/>
        /// </summary>
        /// <param name="inputFile">The binary Excel file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <param name="storageName">The complete or part of the name from the storage that needs to be saved</param>
        /// <returns></returns>
        private List<string> ExtractFromExcelBinaryFormat(string inputFile, string outputFolder, string storageName)
        {
            var compoundFile = new CompoundFile(inputFile);
            
            var result = new List<string>();

            foreach (var child in compoundFile.RootStorage.Children)
            {
                var childStorage = child as CFStorage;
                if (childStorage == null) continue;
                if (!childStorage.Name.StartsWith(storageName)) continue;

                var extractedFileName = ExtractFromStorageNode(compoundFile, childStorage, outputFolder);
                if (extractedFileName != null)
                    result.Add(extractedFileName);

            }

            return result;
        }
        #endregion

        #region ExtractFromPowerPointBinaryFormat
        /// <summary>
        /// This method saves all the PowerPoint embedded binary objects from the <see cref="inputFile"/> to the
        /// <see cref="outputFolder"/>
        /// </summary>
        /// <param name="inputFile">The binary PowerPoint file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns></returns>
        private List<string> ExtractFromPowerPointBinaryFormat(string inputFile, string outputFolder)
        {
            var compoundFile = new CompoundFile(inputFile);

            var result = new List<string>();

            if (compoundFile.RootStorage.ExistsStream("PowerPoint Document"))
            {
                var stream = compoundFile.RootStorage.GetStream("PowerPoint Document") as CFStream;
                var memoryStream = new MemoryStream(stream.GetData());
                using (var binaryReader = new BinaryReader(memoryStream))
                {
                    while (binaryReader.BaseStream.Position != memoryStream.Length)
                    {
                        var verAndInstance = binaryReader.ReadUInt16();
                        var version = verAndInstance & 0x000FU;         // first 4 bit of field verAndInstance
                        var instance = (verAndInstance & 0xFFF0U) >> 4; // last 12 bit of field verAndInstance
                        
                        var typeCode = binaryReader.ReadUInt16();
                        var size = binaryReader.ReadUInt32();
                        var isContainer = (version == 0xF);

                        // Embedded OLE objects start with code 4045
                        if (typeCode == 4113)
                        {
                            if (instance == 0)
                            {
                                // Uncompressed
                                File.WriteAllBytes("d:\\keesje.bin", binaryReader.ReadBytes((int)size));
                            }
                            else
                            {
                                var decompressedSize = binaryReader.ReadUInt32();
                                var data = binaryReader.ReadBytes((int) size - 4);
                                var compressedMemoryStream = new MemoryStream(data);

                                // Skip the first 2 bytes
                                memoryStream.ReadByte();
                                memoryStream.ReadByte();

                                // Decompress the bytes
                                var decompressedBytes = new byte[decompressedSize];
                                var deflateStream = new DeflateStream(compressedMemoryStream, CompressionMode.Decompress, true);
                                deflateStream.Read(decompressedBytes, 0, decompressedBytes.Length);
                                File.WriteAllBytes("d:\\keesje.bin", decompressedBytes);
                            }
                        }
                        else
                        {
                            binaryReader.BaseStream.Position += size;
                        }
                    }
                }
            }


            return result;
        }
        #endregion

 
        #region DecompressPowerPointData
        private byte[] DecompressPowerPointOleData(byte[] data)
        {
            // http://www.idea2ic.com/File_Formats/PowerPoint%2097%20File%20Format.pdf
           
            //deflateStream.Read(decompressedBytes, 0, decompressedBytes.Length);

            //return decompressedBytes;
            return null;
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

        #region ExtractFromStorageNode
        /// <summary>
        /// This method will extract and save the data from the given <see cref="storage"/> node to the <see cref="outputFolder"/>
        /// </summary>
        /// <param name="compoundFile">The <see cref="CompoundFile"/></param>
        /// <param name="storage">The <see cref="CFStorage"/> node</param>
        /// <param name="outputFolder">The outputFolder</param>
        /// <returns></returns>
        private string ExtractFromStorageNode(CompoundFile compoundFile, CFStorage storage, string outputFolder)
        {
            // Embedded objects can be stored in 4 ways
            // - As a CONTENT stream
            // - As a Package
            // - As an Ole10Native object
            // - Embedded into the same compound file
            if (storage.ExistsStream("CONTENTS"))
            {
                var contents = storage.GetStream("CONTENTS");
                if (contents.Size > 0)
                    return SaveByteArrayToFile(contents.GetData(), outputFolder + "Embedded object");
            }
            else if (storage.ExistsStream("Package"))
            {
                var package = storage.GetStream("Package");
                if (package.Size > 0)
                    return SaveByteArrayToFile(package.GetData(), outputFolder + "Embedded object");
            }
            else if (storage.ExistsStream("\x01Ole10Native"))
            {
                var ole10Native = storage.GetStream("\x01Ole10Native");
                if (ole10Native.Size > 0)
                    return ExtractFileFromOle10Native(ole10Native.GetData(), outputFolder);
            }
            else if (storage.ExistsStream("WordDocument"))
            {
                // The embedded object is a Word file
                var tempFileName = FileManager.FileExistsMakeNew(outputFolder + "Embedded Word document.doc");
                compoundFile.SaveStorageTreeToCompoundFile(storage, tempFileName);
                return tempFileName;
            }
            else if (storage.ExistsStream("Workbook"))
            {
                // The embedded object is an Excel file   
                var tempFileName = FileManager.FileExistsMakeNew(outputFolder + "Embedded Excel document.xls");
                compoundFile.SaveStorageTreeToCompoundFile(storage, tempFileName);
                return tempFileName;
            }
            else if (storage.ExistsStream("PowerPoint Document"))
            {
                // The embedded object is a PowerPoint file
                var tempFileName = outputFolder + FileManager.FileExistsMakeNew("Embedded PowerPoint document.ppt");
                compoundFile.SaveStorageTreeToCompoundFile(storage, tempFileName);
                return tempFileName;
            }

            return null;
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
