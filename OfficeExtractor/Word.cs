using System.Collections.Generic;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;
using OfficeExtractor.Ole;
using OpenMcdf;

//
// Word.cs
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
    /// This class is used as a placeholder for all Word related methods
    /// </summary>
    internal class Word
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
        /// This method saves all the Word embedded binary objects from the <paramref name="inputFile"/> to the
        /// <paramref name="outputFolder"/>
        /// </summary>
        /// <param name="inputFile">The binary Word file</param>
        /// <param name="outputFolder">The output folder</param>
        /// <returns></returns>
        /// <exception cref="OEFileIsPasswordProtected">Raised when the <paramref name="inputFile"/> is password protected</exception>
        internal List<string> Extract(string inputFile, string outputFolder)
        {
            using (var compoundFile = new CompoundFile(inputFile))
            {
                var result = new List<string>();

                if(!compoundFile.RootStorage.TryGetStorage("ObjectPool", out var objectPoolStorage))
                    return result;

                Logger.WriteToLog("Object Pool stream found (Word)");
                
                void Entries(CFItem item)
                {
                    var childStorage = item as CFStorage;
                    if (childStorage == null) return;

                    string extractedFileName;

                    if (!childStorage.TryGetStream("\x0001Ole10Native", out _))
                    {
                        if(childStorage.TryGetStream("\x0001CompObj", out var compObj))
                        {
                            var compObjStream = new CompObjStream(compObj);
                            if (compObjStream.AnsiUserType == "OLE Package")
                            {
                                extractedFileName = Extraction.SaveFromStorageNode(childStorage, outputFolder, null);
                                if (!string.IsNullOrEmpty(extractedFileName)) result.Add(extractedFileName);
                                return;
                            }
                        }

                        if(childStorage.TryGetStream("\x0003ObjInfo", out var objInfo))
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

                    if (!string.IsNullOrEmpty(extractedFileName)) result.Add(extractedFileName);
                }

                objectPoolStorage.VisitEntries(Entries, false);

                return result;
            }
        }
        #endregion
    }
}
