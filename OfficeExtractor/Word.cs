using System;
using System.Collections.Generic;
using System.IO;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;
using OfficeExtractor.Ole;
using OpenMcdf;

//
// Word.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2019 Magic-Sessions. (www.magic-sessions.com)
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
                var result = new List<string>();

                var objectPoolStorage = compoundFile.RootStorage.TryGetStorage("ObjectPool");
                if (objectPoolStorage == null) return result;

                Logger.WriteToLog("Object Pool stream found (Word)");
                
                void Entries(CFItem item)
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
                                if (!string.IsNullOrEmpty(extractedFileName)) result.Add(extractedFileName);
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

                    if (!string.IsNullOrEmpty(extractedFileName)) result.Add(extractedFileName);
                }

                objectPoolStorage.VisitEntries(Entries, false);

                return result;
            }
        }
        #endregion
    }
}
