using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using CompoundFileStorage;
using OfficeExtractor.Exceptions;
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
                                        case "\x01Ole10Native":
                                            result.Add(ExtractOleObjectV20(stream, outputFolder));
                                            break;

                                        case "Outlook.FileAttach":
                                            result.Add(ExtractOutlookFileAttachObject(stream, outputFolder));
                                            break;

                                        default:
                                            result.Add(ExtractDefaultObject(stream, outputFolder));
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

        #region ExtractOleObjectV10
        /// <summary>
        /// Extracts a OLE v1.0 object from the given <paramref name="stream"/>
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="outputFolder">The output folder</param>
        internal static string ExtractOleObjectV10(Stream stream, string outputFolder)
        {
            var oleObjectV10 = new ObjectV10(stream);
            var fileName = Path.Combine(outputFolder, oleObjectV10.ItemName);
            return Extraction.SaveByteArrayToFile(oleObjectV10.NativeData, fileName);
        }
        #endregion

        #region ExtractOleObjectV20
        /// <summary>
        /// Extracts a OLE v2.0 object from the given <paramref name="stream"/>
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="outputFolder">The output folder</param>
        internal static string ExtractOleObjectV20(Stream stream, string outputFolder)
        {
            var oleObjectV20 = new ObjectV20(stream);
            var outputFile = Path.Combine(outputFolder, oleObjectV20.FileName ?? "Embedded object");
            return Extraction.SaveByteArrayToFile(oleObjectV20.Data, outputFile);
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

            var bytes = stream.GetData();
            using (var memoryStream = new MemoryStream(bytes))
            using (var binaryReader = new BinaryReader(memoryStream))
            {
                var version = binaryReader.ReadUInt16();
                var longPathName = Strings.Read1ByteLengthPrefixedString(binaryReader);
                var pathName = Strings.Read1ByteLengthPrefixedString(binaryReader);
                var displayName = Strings.Read1ByteLengthPrefixedString(binaryReader);
                var longFileName = Strings.Read1ByteLengthPrefixedString(binaryReader);
                var fileName = Strings.Read1ByteLengthPrefixedString(binaryReader);

                if (!string.IsNullOrEmpty(longFileName)) return longFileName;
                if (!string.IsNullOrEmpty(displayName)) return displayName;
                if (!string.IsNullOrEmpty(fileName)) return fileName;
            }

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
            var oleObjectV10 = new ObjectV10(stream);

            // After that it is wrapped in a compound document
            using (var internalStream = new MemoryStream(oleObjectV10.NativeData))
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

                if (!compoundFile.RootStorage.ExistsStream("AttachContents")) return null;
                
                fileName = Path.Combine(outputFolder, fileName);
                var data = compoundFile.RootStorage.GetStream("AttachContents").GetData();
                return Extraction.SaveByteArrayToFile(data, fileName);
            }
        }
        #endregion

        #region ExtractDefaultObject
        /// <summary>
        /// Extracts a default object from the given <paramref name="stream"/>
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="outputFolder">The output folder</param>
        internal static string ExtractDefaultObject(Stream stream, string outputFolder)
        {
            using (var binaryReader = new BinaryReader(stream))
            {
                var type = binaryReader.ReadByte();
                stream.Position -= 1;

                switch (type)
                {
                    // OLE v1.0
                    case 1:
                        return ExtractOleObjectV10(stream, outputFolder);

                    // OLE v2.0
                    case 2:
                        return ExtractOleObjectV20(stream, outputFolder);

                    default:
                        throw new OEFileIsCorrupt("Invalid embedded object type found '" + type + "', expected 1 or 2");
                }

            }
        }
        #endregion
    }
}
