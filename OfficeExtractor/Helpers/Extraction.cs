using System;
using System.IO;
using System.Text;
using CompoundFileStorage;
using DocumentServices.Modules.Extractors.OfficeExtractor.Exceptions;
using ICSharpCode.SharpZipLib.Zip;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.Helpers
{
    /// <summary>
    /// This class contain helpers method for extraction
    /// </summary>
    internal static class Extraction
    {
        #region GetBytesFromCompoundPackageStream
        /// <summary>
        /// Checks if the <paramref name="bytes"/> is a compound file and if so then tries to extract
        /// the package stream from it. If it fails it will return the original <paramref name="bytes"/>
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns></returns>
        internal static byte[] GetBytesFromCompoundPackageStream(byte[] bytes)
        {
            try
            {
                using (var memoryStream = new MemoryStream(bytes))
                using (var compoundFile = new CompoundFile(memoryStream))
                {
                    if (!compoundFile.RootStorage.ExistsStream("Package"))
                        return bytes;

                    var package = compoundFile.RootStorage.GetStream("Package");
                    return package.GetData();
                }
            }
            catch (Exception)
            {
                return bytes;
            }
        }
        #endregion

        #region GetFileNameFromObjectReplacementFile
        /// <summary>
        /// Tries to extracts the original filename for the OLE object out of the ObjectReplacement file
        /// </summary>
        /// <param name="zipFile"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        internal static string GetFileNameFromObjectReplacementFile(ZipFile zipFile, int index)
        {
            try
            {
                var zipEntry = zipFile[index];
                using (var zipEntryStream = zipFile.GetInputStream(zipEntry))
                using (var zipEntryMemoryStream = new MemoryStream())
                {
                    zipEntryStream.CopyTo(zipEntryMemoryStream);
                    zipEntryMemoryStream.Position = 0x4470;
                    using (var binaryReader = new BinaryReader(zipEntryMemoryStream))
                    {
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
                                return fileName;
                            }
                        }
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

        #region SaveFileFromOle10Native
        /// <summary>
        /// Extract the file from the Ole10Native container and saves it to the outputfolder
        /// </summary>
        /// <param name="ole10Native">The Ole10Native object as an byte array</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>The filename with path from the extracted file</returns>
        internal static string SaveFileFromOle10Native(byte[] ole10Native, string outputFolder)
        {
            // Convert the byte array to a stream
            using (Stream oleStream = new MemoryStream(ole10Native))
            {
                // The name of the file start at postion 7 so move to there
                oleStream.Position = 6;
                var fileName = string.Empty;

                // Read until we find a null character
                int i;
                var chr = new byte[1];
                for (i = 0; i < 260; i++)
                {
                    oleStream.Read(chr, 0, 1);
                    if (chr[0] == 0)
                        break;

                    // Unicode char found
                    if (chr[0] >= 0xc2 && chr[0] <= 0xdf)
                    {
                        i += 1;

                        var chr2 = new byte[2];
                        chr2[1] = chr[0];
                        oleStream.Read(chr, 0, 1);
                        chr2[0] = chr[0];

                        fileName += Encoding.GetEncoding("ANSI6").GetString(chr2);
                    }
                    else
                        fileName += (char)chr[0];
                }

                // We don't need this but we need to read it to know where we
                // are located in the stream
                var tempOriginalFilePath = new char[260];
                for (i = 0; i < 260; i++)
                {
                    oleStream.Read(chr, 0, 1);
                    tempOriginalFilePath[i] = (char)chr[0];
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

        #region SaveFromStorageNode
        /// <summary>
        /// This method will extract and save the data from the given <see cref="storage"/> node to the <see cref="outputFolder"/>
        /// </summary>
        /// <param name="storage">The <see cref="CFStorage"/> node</param>
        /// <param name="outputFolder">The outputFolder</param>
        /// <returns></returns>
        /// <exception cref="OEFileIsPasswordProtected">Raised when a WordDocument, WorkBook or PowerPoint Document stream is password protected</exception>
        internal static string SaveFromStorageNode(CFStorage storage, string outputFolder)
        {
            return SaveFromStorageNode(storage, outputFolder, null);
        }

        /// <summary>
        /// This method will extract and save the data from the given <see cref="storage"/> node to the <see cref="outputFolder"/>
        /// </summary>
        /// <param name="storage">The <see cref="CFStorage"/> node</param>
        /// <param name="outputFolder">The outputFolder</param>
        /// <param name="fileName">The fileName to use, null when the fileName is unknown</param>
        /// <returns></returns>
        /// <exception cref="OEFileIsPasswordProtected">Raised when a WordDocument, WorkBook or PowerPoint Document stream is password protected</exception>
        public static string SaveFromStorageNode(CFStorage storage, string outputFolder, string fileName)
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
                    return SaveByteArrayToFile(contents.GetData(), outputFolder + (fileName ?? "Embedded object"));
            }
            else if (storage.ExistsStream("Package"))
            {
                var package = storage.GetStream("Package");
                if (package.Size > 0)
                    return SaveByteArrayToFile(package.GetData(), outputFolder + (fileName ?? "Embedded object"));
            }
            else if (storage.ExistsStream("\x01Ole10Native"))
            {
                var ole10Native = storage.GetStream("\x01Ole10Native");
                if (ole10Native.Size > 0)
                    return SaveFileFromOle10Native(ole10Native.GetData(), outputFolder);
            }
            else if (storage.ExistsStream("WordDocument"))
            {
                // The embedded object is a Word file
                var tempFileName = outputFolder + (fileName ?? FileManager.FileExistsMakeNew("Embedded Word document.doc"));
                SaveStorageTreeToCompoundFile(storage, tempFileName);
                return tempFileName;
            }
            else if (storage.ExistsStream("Workbook"))
            {
                // The embedded object is an Excel file   
                var tempFileName = outputFolder + (fileName ?? FileManager.FileExistsMakeNew("Embedded Excel document.xls"));
                SaveStorageTreeToCompoundFile(storage, tempFileName);
                return tempFileName;
            }
            else if (storage.ExistsStream("PowerPoint Document"))
            {
                // The embedded object is a PowerPoint file
                var tempFileName = outputFolder + (fileName ?? FileManager.FileExistsMakeNew("Embedded PowerPoint document.ppt"));
                SaveStorageTreeToCompoundFile(storage, tempFileName);
                return tempFileName;
            }

            return null;
        }
        #endregion

        #region SaveStorageTreeToCompoundFile
        /// <summary>
        /// This will save the complete tree from the given <paramref name="storage"/> to a new <see cref="CompoundFile"/>
        /// </summary>
        /// <param name="storage"></param>
        /// <param name="fileName">The filename with path for the new compound file</param>
        internal static void SaveStorageTreeToCompoundFile(CFStorage storage, string fileName)
        {
            using (var compoundFile = new CompoundFile())
            {
                GetStorageChain(compoundFile.RootStorage, storage);
                var extension = Path.GetExtension(fileName);

                if (extension != null)
                    switch (extension.ToUpperInvariant())
                    {
                        case ".XLS":
                        case ".XLT":
                        case ".XLW":
                            Excel.SetWorkbookVisibility(compoundFile);
                            break;
                    }

                compoundFile.Save(fileName);
            }
        }

        /// <summary>
        /// Returns the complete tree with all the <see cref="CFStorage"/> and <see cref="CFStream"/> children
        /// </summary>
        /// <param name="rootStorage"></param>
        /// <param name="storage"></param>
        private static void GetStorageChain(CFStorage rootStorage, CFStorage storage)
        {
            foreach (var child in storage.Children)
            {
                if (child.IsStorage)
                {
                    var newRootStorage = rootStorage.AddStorage(child.Name);
                    GetStorageChain(newRootStorage, child as CFStorage);
                }
                else if (child.IsStream)
                {
                    var childStream = child as CFStream;
                    if (childStream == null) continue;
                    var stream = rootStorage.AddStream(child.Name);
                    var bytes = childStream.GetData();
                    stream.SetData(bytes);
                }
            }
        }
        #endregion

        #region SaveByteArrayToFile
        /// <summary>
        /// Saves the <paramref name="data"/> byte array to the <paramref name="outputFile"/>
        /// </summary>
        /// <param name="data">The stream as byte array</param>
        /// <param name="outputFile">The output filename with path</param>
        /// <returns></returns>
        /// <exception cref="OEFileIsCorrupt">Raised when the file is corrupt</exception> 
        internal static string SaveByteArrayToFile(byte[] data, string outputFile)
        {
            // Because the data is stored in a stream we have no name for it so we
            // have to check the magic bytes to see with what kind of file we are dealing

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

            if (extension != null)
            {
                switch (extension.ToUpperInvariant())
                {
                    case ".XLS":
                    case ".XLT":
                    case ".XLW":
                        using (var memoryStream = new MemoryStream(data))
                        using (var compoundFile = new CompoundFile(memoryStream))
                        {
                            Excel.SetWorkbookVisibility(compoundFile);
                            compoundFile.Save(outputFile);
                        }
                        break;

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
            }
            else
                File.WriteAllBytes(outputFile, data);

            return outputFile;
        }
        #endregion
    }
}
