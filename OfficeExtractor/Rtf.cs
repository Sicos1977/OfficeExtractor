using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using CompoundFileStorage;
using OfficeExtractor.Helpers;
using OfficeExtractor.Ole;
using Path = System.IO.Path;

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
                                        case "MailMsgAtt":
                                        case "Outlook.FileAttach":
                                            result.Add(ExtractOutlookFileAttachObject(stream, outputFolder));
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
            var outputFile = string.IsNullOrWhiteSpace(ole10.ItemName)
                    ? Extraction.DefaultEmbeddedObjectName
                    : ole10.ItemName;

            if (ole10.Format == OleFormat.File)
                return Extraction.IsCompoundFile(ole10.NativeData)
                    ? Extraction.SaveFromStorageNode(ole10.NativeData, outputFolder, ole10.ItemName)
                    : Extraction.SaveByteArrayToFile(ole10.NativeData, Path.Combine(outputFolder, outputFile));

            return null;
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
            if (!string.IsNullOrEmpty(ad.LongFileName)) return ad.LongFileName;
            if (!string.IsNullOrEmpty(ad.DisplayName)) return ad.DisplayName;
            if (!string.IsNullOrEmpty(ad.FileName)) return ad.FileName;
            return null;
        }
        #endregion

        #region ExtractOutlookFileAttachObject
        /// <summary>
        /// Extracts a Outlook File Attachment object from the given <paramref name="stream"/>
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="outputFolder">The output folder</param>
        internal static string ExtractOutlookFileAttachObject(Stream stream, string outputFolder)
        {
            // Outlook attachments embedded in RTF are firstly embedded in an OLE v1.0 object
            var ole10 = new Ole10(stream);

            // After that it is wrapped in a compound document
            using (var internalStream = new MemoryStream(ole10.NativeData))
            using (var compoundFile = new CompoundFile(internalStream))
            {
                string fileName = null;
                File.WriteAllBytes("d:\\kees.txt", ole10.NativeData);
                if (compoundFile.RootStorage.ExistsStream("AttachDesc"))
                {
                    var attachDescStream = compoundFile.RootStorage.GetStream("AttachDesc") as CFStream;
                    fileName = GetFileNameFromAttachDescStream(attachDescStream);
                }

                if (string.IsNullOrEmpty(fileName))
                    fileName = Extraction.DefaultEmbeddedObjectName;

                if (!compoundFile.RootStorage.ExistsStream("AttachContents")) return null;
                
                fileName = Path.Combine(outputFolder, fileName);
                var data = compoundFile.RootStorage.GetStream("AttachContents").GetData();
                return Extraction.SaveByteArrayToFile(data, fileName);
            }
        }
        #endregion
    }
}
