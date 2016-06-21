using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;
using CompoundFileStorage;
using ICSharpCode.SharpZipLib.Zip;

/*
   Copyright 2014-2016 Kees van Spelde

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
    /// This class is used to extract embedded files from Word, Excel and PowerPoint files. It only extracts
    /// one level deep, for example when you have an Word file with an embedded Excel file that has an embedded
    /// PDF it will only extract the embedded Excel file from the Word file.
    /// </summary>
    public class Extractor
    {
        #region CheckFileNameAndOutputFolder
        /// <summary>
        /// Checks if the <see cref="inputFile"/> and <see cref="outputFolder"/> is valid
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
                        // Aan de eerste 128 bytes hebben we genoeg om de bestandstypes te herkennen
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

        #region SaveToFolder
        /// <summary>
        /// Extracts all the embedded object from the Microsoft Office <paramref name="inputFile"/> to the 
        /// <see cref="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="inputFile">The Microsoft Office file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        /// <exception cref="ArgumentNullException">Raised when the <paramref name="inputFile"/> or <paramref name="outputFolder"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <sparamref name="inputFile"/> does not exist</exception>
        /// <exception cref="DirectoryNotFoundException">Raised when the <paramref name="outputFolder"/> does not exists</exception>
        /// <exception cref="OEFileIsCorrupt">Raised when the <paramref name="inputFile" /> is corrupt</exception>
        /// <exception cref="OEFileTypeNotSupported">Raised when the <paramref name="inputFile"/> is not supported</exception>
        /// <exception cref="OEFileIsPasswordProtected">Raised when the <paramref name="inputFile"/> is password protected</exception>
        public List<string> SaveToFolder(string inputFile, string outputFolder)
        {
            CheckFileNameAndOutputFolder(inputFile, outputFolder);

            var extension = GetExtension(inputFile);

            outputFolder = FileManager.CheckForBackSlash(outputFolder);

            switch (extension)
            {
                case ".ODT":
                case ".ODS":
                case ".ODP":
                    return ExtractFromOpenDocumentFormat(inputFile, outputFolder);

                case ".DOC":
                case ".DOT":
                    // Word 97 - 2003
                    return Word.SaveToFolder(inputFile, outputFolder);

                case ".DOCM":
                case ".DOCX":
                case ".DOTM":
                    // Word 2007 - 2013
                    return ExtractFromOfficeOpenXmlFormat(inputFile, "/word/embeddings/", outputFolder);

                case ".RTF":
                    return Rtf.SaveToFolder(inputFile, outputFolder);

                case ".XLS":
                case ".XLT":
                case ".XLW":
                    // Excel 97 - 2003
                    return Excel.SaveToFolder(inputFile, outputFolder);

                case ".XLSB":
                case ".XLSM":
                case ".XLSX":
                case ".XLTM":
                case ".XLTX":
                    // Excel 2007 - 2013
                    return ExtractFromOfficeOpenXmlFormat(inputFile, "/xl/embeddings/", outputFolder);

                case ".POT":
                case ".PPT":
                case ".PPS":
                    // PowerPoint 97 - 2003
                    return PowerPoint.SaveToFolder(inputFile, outputFolder);

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
                                                     "' is not supported, only .ODT, .DOC, .DOCM, .DOCX, .DOT, .DOTM, .RTF, .XLS, .XLSB, .XLSM, .XLSX, .XLT, " +
                                                     ".XLTM, .XLTX, .XLW, .POT, .PPT, .POTM, .POTX, .PPS, .PPSM, .PPSX, .PPTM and .PPTX are supported");
            }
        }
        #endregion

        #region ExtractFromOfficeOpenXmlFormat
        /// <summary>
        /// Extracts all the embedded object from the Office Open XML <paramref name="inputFile"/> to the 
        /// <see cref="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="inputFile">The Office Open XML format file</param>
        /// <param name="embeddingPartString">The folder in the Office Open XML format (zip) file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        /// <exception cref="OEFileIsPasswordProtected">Raised when the Microsoft Office file is password protected</exception>
        internal List<string> ExtractFromOfficeOpenXmlFormat(string inputFile, string embeddingPartString, string outputFolder)
        {
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
                                    using (var compoundFile = new CompoundFile(packagePartStream))
                                    {
                                        result.Add(Extraction.SaveFromStorageNode(compoundFile.RootStorage, outputFolder));
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
                    if (
                        !fileFormatException.Message.Equals("File contains corrupted data.",
                            StringComparison.InvariantCultureIgnoreCase))
                        return null;

                    try
                    {
                        // When we receive this exception we can have 2 things:
                        // - The file is corrupt
                        // - The file is password protected, in this case the file is saved as a compound file
                        //EncryptedPackage
                        using (var compoundFile = new CompoundFile(inputFileMemoryStream))
                        {
                            if (compoundFile.RootStorage.ExistsStream("EncryptedPackage"))
                                throw new OEFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                                    "' is password protected");
                        }

                    }
                    catch (Exception)
                    {
                        return null;
                    }
                }
            }

            return null;
        }
        #endregion

        #region ExtractFromOpenDocumentFormat
        /// <summary>
        /// Extracts all the embedded object from the OpenDocument <paramref name="inputFile"/> to the 
        /// <see cref="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="inputFile">The OpenDocument format file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        /// <exception cref="OEFileIsPasswordProtected">Raised when the OpenDocument format file is password protected</exception>
        internal List<string> ExtractFromOpenDocumentFormat(string inputFile, string outputFolder)
        {
            var result = new List<string>();

            var zipFile = new ZipFile(inputFile);
  
            // Check if the file is password protected
            var manifestEntry = zipFile.FindEntry("META-INF/manifest.xml", true);
            if (manifestEntry != -1)
            {
                using (var manifestEntryStream = zipFile.GetInputStream(manifestEntry))
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

            foreach (ZipEntry zipEntry in zipFile)
            {
                if (!zipEntry.IsFile) continue;
                if (zipEntry.IsCrypted)
                    throw new OEFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                                "' is password protected");

                var name = zipEntry.Name.ToUpperInvariant();
                if (!name.StartsWith("OBJECT") || name.Contains("/"))
                    continue;

                string fileName = null;

                var objectReplacementFileIndex = zipFile.FindEntry("ObjectReplacements/" + name, true);
                if (objectReplacementFileIndex != -1)
                    fileName = Extraction.GetFileNameFromObjectReplacementFile(zipFile, objectReplacementFileIndex);
                
                using (var zipEntryStream = zipFile.GetInputStream(zipEntry))
                using (var zipEntryMemoryStream = new MemoryStream())
                {
                    zipEntryStream.CopyTo(zipEntryMemoryStream);

                    using (var compoundFile = new CompoundFile(zipEntryMemoryStream))
                        result.Add(Extraction.SaveFromStorageNode(compoundFile.RootStorage, outputFolder, fileName));
                }
            }

            return result;
        }
        #endregion
    }
}
