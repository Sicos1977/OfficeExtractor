using System.Collections.Generic;
using System.IO;
using CompoundFileStorage;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;

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
            using (var compoundFile = new CompoundFile(inputFile))
            {
                if (IsPasswordProtected(compoundFile))
                    throw new OEFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                        "' is password protected");

                var result = new List<string>();

                if (!compoundFile.RootStorage.ExistsStorage("ObjectPool")) return result;
                var objectPoolStorage = compoundFile.RootStorage.GetStorage("ObjectPool") as CFStorage;
                if (objectPoolStorage == null) return result;

                // Multiple objects are stored as children of the storage object
                foreach (var child in objectPoolStorage.Children)
                {
                    var childStorage = child as CFStorage;
                    if (childStorage == null) continue;
                    var extractedFileName = Extraction.SaveFromStorageNode(childStorage, outputFolder);
                    if (extractedFileName != null)
                        result.Add(extractedFileName);
                }

                return result;
            }
        }
        #endregion

        #region IsPasswordProtected
        /// <summary>
        /// Returns true when the Word file is password protected
        /// </summary>
        /// <param name="compoundFile"></param>
        /// <returns></returns>
        /// <exception cref="OEFileIsCorrupt">Raised when the file is corrupt</exception>
        public static bool IsPasswordProtected(CompoundFile compoundFile)
        {
            if (!compoundFile.RootStorage.ExistsStream("WordDocument")) 
                throw new OEFileIsCorrupt("Could not find the WordDocument stream in the file '" + compoundFile.FileName + "'");
            var stream = compoundFile.RootStorage.GetStream("WordDocument") as CFStream;
            if (stream == null) return false;

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
