using System.Collections.Generic;
using System.IO;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;
using OfficeExtractor.Ole;
using OpenMcdf;
using Path = System.IO.Path;

//
// Rtf.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2020 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

namespace OfficeExtractor
{
    /// <summary>
    /// This class is used as a placeholder for all RTF related methods
    /// </summary>
    internal class Rtf
    {
        #region Fields
        /// <summary>
        ///     <see cref="Extraction"/>
        /// </summary>
        private Extraction _extraction;
        #endregion

        #region Properties
        /// <summary>
        /// Returns a reference to the Extraction class when it already exists or creates a new one
        /// when it doesn't
        /// </summary>
        private Extraction Extraction
        {
            get
            {
                if (_extraction != null)
                    return _extraction;

                _extraction = new Extraction();
                return _extraction;
            }
        }
        #endregion

        #region Extract
        /// <summary>
        /// Saves all the embedded object from the RTF <paramref name="inputFile"/> to the 
        /// <paramref name="outputFolder"/> and returns the files with full path as a list of strings
        /// </summary>
        /// <param name="inputFile">The RTF file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        internal List<string> Extract(string inputFile, string outputFolder)
        {
            var result = new List<string>();

            using (var streamReader = new StreamReader(inputFile))
            {
                var rtfReader = new RtfParser.Reader(streamReader);
                var enumerator = rtfReader.Read().GetEnumerator();
                while (enumerator.MoveNext())
                {
                    if (enumerator.Current?.Text != "object") continue;
                    if (!RtfParser.Reader.MoveToNextControlWord(enumerator, "objclass")) continue;
                    var className = RtfParser.Reader.GetNextText(enumerator);

                    if (!RtfParser.Reader.MoveToNextControlWord(enumerator, "objdata")) continue;
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
            return result;
        }
        #endregion

        #region ExtractOle10
        /// <summary>
        /// Extracts a OLE v1.0 object from the given <paramref name="stream"/>
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="outputFolder">The output folder</param>
        internal string ExtractOle10(Stream stream, string outputFolder)
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

                case "PBrush":
                    // Ignore
                    return null;

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
        private string GetFileNameFromAttachDescStream(CFStream stream)
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
        internal string ExtractOutlookAttachmentObject(Stream stream, string outputFolder)
        {
            // Outlook attachments embedded in RTF are firstly embedded in an OLE v1.0 object
            var ole10 = new Ole10(stream);

            // After that it is wrapped in a compound document
            using (var internalStream = new MemoryStream(ole10.NativeData))
            using (var compoundFile = new CompoundFile(internalStream))
            {
                string fileName = null;
                if(compoundFile.RootStorage.TryGetStream("AttachDesc", out var attachDescStream))
                    fileName = GetFileNameFromAttachDescStream(attachDescStream);

                if (string.IsNullOrEmpty(fileName))
                    fileName = Extraction.DefaultEmbeddedObjectName;

                fileName = FileManager.RemoveInvalidFileNameChars(fileName);
                fileName = Path.Combine(outputFolder, fileName);
                fileName = FileManager.FileExistsMakeNew(fileName);

                if (compoundFile.RootStorage.TryGetStream("AttachContents", out var attachContentsStream))
                    return Extraction.SaveByteArrayToFile(attachContentsStream.GetData(), fileName);

                if(compoundFile.RootStorage.TryGetStorage("MAPIMessage", out var mapiMessageStorage))
                {
                    fileName = Path.Combine(outputFolder, fileName);
                    return Extraction.SaveStorageTreeToCompoundFile(mapiMessageStorage, fileName);
                }

                return null;
            }
        }
        #endregion
    }
}
