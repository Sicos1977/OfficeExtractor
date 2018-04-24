using System;
using System.Collections.Generic;
using System.IO;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;
using OfficeExtractor.Ole;
using OpenMcdf;

/*
   Copyright 2013 - 2018 Kees van Spelde

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
    /// This class is used as a placeholder for all Word related methods
    /// </summary>
    internal static class Word
    {
        #region SaveToFolder
        /// <summary>
        /// This method saves all the Word embedded binary objects from the <paramref name="inputFile"/> to the
        /// <see cref="outputFolder"/>
        /// </summary>
        /// <param name="inputFile">The binary Word file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns></returns>
        /// <exception cref="OEFileIsPasswordProtected">Raised when the <see cref="inputFile"/> is password protected</exception>
        public static List<string> SaveToFolder(string inputFile, string outputFolder)
        {
            var fileName = Path.GetFileName(inputFile);

            using (var compoundFile = new CompoundFile(inputFile))
            {
                if (IsPasswordProtected(compoundFile, fileName))
                    throw new OEFileIsPasswordProtected("The file '" + fileName +
                                                        "' is password protected");

                var result = new List<string>();

                var objectPoolStorage = compoundFile.RootStorage.TryGetStorage("ObjectPool");
                if (objectPoolStorage == null) return result;

                Action<CFItem> entries = item =>
                {
                    var childStorage = item as CFStorage;
                    if (childStorage == null) return;

                    string extractedFileName;

                    if (childStorage.TryGetStream("\x0001Ole10Native") != null)
                    {
                        var compObj = childStorage.TryGetStream("\x0001CompObj");
                        if (compObj != null)
                        {
                            var compObjStream = new CompObjStream(compObj);
                            if (compObjStream.AnsiUserType == "OLE Package")
                            {
                                extractedFileName = Extraction.SaveFromStorageNode(childStorage, outputFolder, null);
                                if (!string.IsNullOrEmpty(extractedFileName))
                                    result.Add(extractedFileName);
                                return;
                            }
                        }
                        
                        var objInfo = childStorage.TryGetStream("\x0003ObjInfo");
                        if (objInfo != null)
                        {
                            var objInfoStream = new ObjInfoStream(objInfo);
                            // We don't want to export linked objects and objects that are not shown as an icon... 
                            // because these objects are already visible on the Word document
                            if (objInfoStream.Link || !objInfoStream.Icon) return;
                        }

                        extractedFileName = Extraction.SaveFromStorageNode(childStorage, outputFolder, null);
                    }
                    else
                    {
                        // Get the objInfo stream to check if this is a linked file... if so then ignore it
                        var objInfo = childStorage.GetStream("\x0003ObjInfo");
                        var objInfoStream = new ObjInfoStream(objInfo);
                        
                        // We don't want to export linked objects and objects that are not shown as an icon... 
                        // because these objects are already visible on the Word document
                        if (objInfoStream.Link || !objInfoStream.Icon) return;
                        extractedFileName = Extraction.SaveFromStorageNode(childStorage, outputFolder);
                    }

                    if (!string.IsNullOrEmpty(extractedFileName))
                        result.Add(extractedFileName);
                };

                objectPoolStorage.VisitEntries(entries, false);

                return result;
            }
        }
        #endregion

        #region IsPasswordProtected
        /// <summary>
        /// Returns true when the Word file is password protected
        /// </summary>
        /// <param name="compoundFile"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        /// <exception cref="OEFileIsCorrupt">Raised when the file is corrupt</exception>
        public static bool IsPasswordProtected(CompoundFile compoundFile, string fileName)
        {
            
            if (compoundFile.RootStorage.TryGetStream("EncryptedPackage") != null) return true;

            var stream = compoundFile.RootStorage.TryGetStream("WordDocument");

            if (stream == null)
                throw new OEFileIsCorrupt("Could not find the WordDocument stream in the file '" + fileName + "'");

            var bytes = stream.GetData();
            using (var memoryStream = new MemoryStream(bytes))
            using (var binaryReader = new BinaryReader(memoryStream))
            {
                //http://msdn.microsoft.com/en-us/library/dd944620%28v=office.12%29.aspx
                // The bit that shows if the file is encrypted is in the 11th and 12th byte so we 
                // need to skip the first 10 bytes
                binaryReader.ReadBytes(10);

                // Now we read the 2 bytes that we need
                var pnNext = binaryReader.ReadUInt16();
                //(value & mask) == mask)

                // The bit that tells us if the file is encrypted
                return (pnNext & 0x0100) == 0x0100;
            }
        }
        #endregion
    }
}
