using System;
using System.IO;
using System.Text;
using CompoundFileStorage;
using OfficeExtractor.Exceptions;
using ICSharpCode.SharpZipLib.Zip;
using OfficeExtractor.Ole;

namespace OfficeExtractor.Helpers
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
                {
                    using (var stream = new MemoryStream(ole10Native.GetData()))
                    {
                        var oleObjectV20 = new ObjectV20(stream);
                        var outputFile = Path.Combine(outputFolder, oleObjectV20.FileName ?? "Embedded object");
                        return SaveByteArrayToFile(oleObjectV20.Data, outputFile);
                    } 
                }
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
