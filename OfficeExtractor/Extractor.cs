using System;
using System.Collections.Generic;
using System.IO;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage;
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

        #region ExtractFromWord
        /// <summary>
        /// Extracts all the embedded Word object from the <see cref="inputFile"/> to the 
        /// <see cref="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="inputFile">The Word file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        /// <exception cref="ArgumentNullException">Raised when the <see cref="inputFile"/> or <see cref="outputFolder"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <see cref="inputFile"/> does not exists</exception>
        /// <exception cref="DirectoryNotFoundException">Raised when the <see cref="outputFolder"/> does not exists</exception>
        public List<string> ExtractFromWord(string inputFile, string outputFolder)
        {
            CheckFileNameAndOutputFolder(inputFile, outputFolder);
            outputFolder = FileManager.CheckForBackSlash(outputFolder);

            // TODO: Add support for Word 2007 and up format

            return ExtractFromWordBinaryFormat(inputFile, outputFolder);
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

                    // Ole objects can be stored in 2 ways
                    // - Directly in the CONTENT stream
                    // - As an Ole10Native object
                    if (childStorage.ExistsStream("CONTENTS"))
                    {
                        var contents = childStorage.GetStream("CONTENTS");
                        // If there is any data
                        if (contents.Size > 0)
                        {
                            var data = contents.GetData();
                            var tempFileName = outputFolder + "Embedded word object";
                            // Because the data is stored in the CONTENT stream we have no name for it so we
                            // have to check the magic bytes to see with what kind of file we are dealing
                            var fileType = FileTypeSelector.GetFileTypeFileInfo(data);
                            if(fileType != null && !string.IsNullOrEmpty(fileType.Extension))
                                tempFileName += "." + fileType.Extension;

                            // Check if the output file already exists and if so make a new one
                            tempFileName = FileManager.FileExistsMakeNew(tempFileName);

                            File.WriteAllBytes(tempFileName, data);
                            result.Add(tempFileName);
                        }
                    }
                    else if (childStorage.ExistsStream("\x01Ole10Native"))
                    {
                        var ole10Native = childStorage.GetStream("\x01Ole10Native");
                        if (ole10Native.Size > 0)
                            result.Add(ExtractFileFromOle10Native(ole10Native.GetData(), outputFolder));
                    }
                }
            }

            return result;
        }
        #endregion

        #region ExtractFromWordOfficeOpenXMLFormat
        /// <summary>
        /// Extracts all the embedded Word object from the <see cref="inputFile"/> to the 
        /// <see cref="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="inputFile">The office open XML Word file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        public List<string> ExtractFromWordOfficeOpenXmlFormat(string inputFile, string outputFolder)
        {
            throw new NotImplementedException();
        }
        #endregion

        #region ExtractFileFromOle10Native
        /// <summary>
        /// Extract the file from the Ole10Native container and saves it to the outputfolder
        /// </summary>
        /// <param name="ole10Native"></param>
        /// <param name="outputFolder"></param>
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
