using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage;
using DocumentServices.Modules.Extractors.OfficeExtractor.Helpers;

namespace DocumentServices.Modules.Extractors.OfficeExtractor
{
    public class Extractor
    {
        #region CheckFileNameAndOutputFolder
        /// <summary>
        /// Checks if the <see cref="fileName"/> and <see cref="outputFolder"/> is valid
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="outputFolder"></param>
        /// <exception cref="ArgumentNullException">Raised when the <see cref="fileName"/> or <see cref="outputFolder"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <see cref="fileName"/> does not exists</exception>
        /// <exception cref="DirectoryNotFoundException">Raised when the <see cref="outputFolder"/> does not exists</exception>
        private void CheckFileNameAndOutputFolder(string fileName, string outputFolder)
        {
            if (string.IsNullOrEmpty(fileName))
                throw new ArgumentNullException(fileName);

            if (string.IsNullOrEmpty(outputFolder))
                throw new ArgumentNullException(outputFolder);

            if (!File.Exists(fileName))
                throw new FileNotFoundException(fileName);

            if (!Directory.Exists(outputFolder))
                throw new DirectoryNotFoundException(outputFolder);
        }
        #endregion

        /// <summary>
        /// Extracts all the embedded Word object from the <see cref="fileName"/> to the 
        /// <see cref="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="fileName">The Word file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        /// <exception cref="ArgumentNullException">Raised when the <see cref="fileName"/> or <see cref="outputFolder"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <see cref="fileName"/> does not exists</exception>
        /// <exception cref="DirectoryNotFoundException">Raised when the <see cref="outputFolder"/> does not exists</exception>
        public List<string> ExtractFromWord(string fileName, string outputFolder)
        {
            CheckFileNameAndOutputFolder(fileName, outputFolder);

            var result = new List<string>();

            var compoundFile = new CompoundFile(fileName);

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
                    // - As an ole10Native object
                    if (childStorage.ExistsStream("CONTENTS"))
                    {
                        var contents = childStorage.GetStream("CONTENTS");
                        // If there is any data
                        if (contents.Size > 0)
                        {
                            var data = contents.GetData();
                            // Because the data is stored in the CONTENT stream we have no name for it so we
                            // have to check the magic bytes to see with what kind of file we are dealing
                            var extension = FileTypeSelector.GetFileTypeFileInfo(data);

                        }
                    }
                    else if (childStorage.ExistsStream("ole10Native"))
                    {
                        var ole10Native = childStorage.GetStream("ole10Native");
                    }
                }
            }
        }

    }
}
