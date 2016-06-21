using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using CompoundFileStorage;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;
using OfficeExtractor.Ole;
using Path = System.IO.Path;

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
    /// This class is used as a placeholder for all RTF related methods
    /// </summary>
    internal static class Rtf
    {
        #region ExtractFromRtf
        /// <summary>
        /// Saves all the embedded object from the RTF <paramref name="inputFile"/> to the 
        /// <see cref="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="inputFile">The RTF file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        internal static List<string> SaveToFolder(string inputFile, string outputFolder)
        {
            var result = new List<string>();

            using (var streamReader = new StreamReader(inputFile))
            {
                var rtfReader = new RtfParser.Reader(streamReader);
                var enumerator = rtfReader.Read().GetEnumerator();
                while (enumerator.MoveNext())
                {
                    if (enumerator.Current.Text == "object")
                    {
                        if (RtfParser.Reader.MoveToNextControlWord(enumerator, "objclass"))
                        {
                            var className = RtfParser.Reader.GetNextText(enumerator);

                            if (RtfParser.Reader.MoveToNextControlWord(enumerator, "objdata"))
                            {
                                var data = RtfParser.Reader.GetNextTextAsByteArray(enumerator);
                                using (var stream = new MemoryStream(data))
                                {

                                    switch (className)
                                    {
                                        case "Outlook.FileAttach":
                                        case "MailMsgAtt":
                                            result.Add(ExtractOutlookAttachmentObject(stream, outputFolder));
                                            break;

                                        default:
                                            var fileName = ExtractOle10(stream, outputFolder);
                                            if (!string.IsNullOrWhiteSpace(fileName))
                                                result.Add(fileName);
                                            break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return result;
        }
        #endregion

        #region ExtractOle10
        /// <summary>
        /// Extracts a OLE v1.0 object from the given <paramref name="stream"/>
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="outputFolder">The output folder</param>
        internal static string ExtractOle10(Stream stream, string outputFolder)
        {
            var ole10 = new Ole10(stream);

            if (ole10.Format != OleFormat.File) return null;
            
            switch (ole10.ClassName)
            {
                case "Package":
                    var package = new Package(ole10.NativeData);
                    if (package.Format == OleFormat.Link) return null;

                    var fileName = Path.GetFileName(package.FileName);
                    if (string.IsNullOrWhiteSpace(fileName))
                        fileName = Extraction.DefaultEmbeddedObjectName;

                    fileName = Path.Combine(outputFolder, fileName);
                    return Extraction.SaveByteArrayToFile(package.Data, fileName);

                default:
                    if (Extraction.IsCompoundFile(ole10.NativeData))
                        return Extraction.SaveFromStorageNode(ole10.NativeData, outputFolder, ole10.ItemName);

                    throw new OEObjectTypeNotSupported("Unsupported OleNative ClassName '" +
                                                       ole10.ClassName + "' found");
            }
        }
        #endregion

        #region GetFileNameFromAttachDescStream
        /// <summary>
        /// Returns the filename from the AttachDesc stream
        /// </summary>
        /// <param name="stream">The AttachDesc stream</param>
        /// <returns></returns>
        [SuppressMessage("ReSharper", "UnusedVariable")]
        private static string GetFileNameFromAttachDescStream(CFStream stream)
        {
            // https://msdn.microsoft.com/en-us/library/ee157577(v=exchg.80).aspx
            if (stream == null) return null;
            var ad = new AttachDescStream(stream);
            
            if (!string.IsNullOrWhiteSpace(ad.LongFileName)) 
                return ad.LongFileName;
            if (!string.IsNullOrWhiteSpace(ad.DisplayName))
                return ad.DisplayName;

            if (!string.IsNullOrWhiteSpace(ad.FileName))
                return ad.FileName;

            return null;
        }
        #endregion

        #region ExtractOutlookAttachmentObject
        /// <summary>
        /// Extracts a Outlook File Attachment object from the given <paramref name="stream"/>
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="outputFolder">The output folder</param>
        internal static string ExtractOutlookAttachmentObject(Stream stream, string outputFolder)
        {
            // Outlook attachments embedded in RTF are firstly embedded in an OLE v1.0 object
            var ole10 = new Ole10(stream);

            // After that it is wrapped in a compound document
            using (var internalStream = new MemoryStream(ole10.NativeData))
            using (var compoundFile = new CompoundFile(internalStream))
            {
                string fileName = null;
                if (compoundFile.RootStorage.ExistsStream("AttachDesc"))
                {
                    var attachDescStream = compoundFile.RootStorage.GetStream("AttachDesc") as CFStream;
                    fileName = GetFileNameFromAttachDescStream(attachDescStream);
                }

                if (string.IsNullOrEmpty(fileName))
                    fileName = Extraction.DefaultEmbeddedObjectName;

                fileName = FileManager.RemoveInvalidFileNameChars(fileName);
                fileName = Path.Combine(outputFolder, fileName);
                fileName = FileManager.FileExistsMakeNew(fileName);

                if (compoundFile.RootStorage.ExistsStream("AttachContents"))
                {
                    var data = compoundFile.RootStorage.GetStream("AttachContents").GetData();
                    return Extraction.SaveByteArrayToFile(data, fileName);
                }

                if (compoundFile.RootStorage.ExistsStorage("MAPIMessage"))
                {

                    fileName = Path.Combine(outputFolder, fileName);
                    var storage = compoundFile.RootStorage.GetStorage("MAPIMessage") as CFStorage;
                    return Extraction.SaveStorageTreeToCompoundFile(storage, fileName);
                }

                return null;
            }
        }
        #endregion
    }
}
