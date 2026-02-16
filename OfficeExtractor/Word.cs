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
// Copyright (c) 2013-2026 Kees van Spelde. (www.magic-sessions.com)
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

namespace OfficeExtractor;

/// <summary>
///     This class is used as a placeholder for all Word related methods
/// </summary>
internal class Word : OfficeBase
{
    #region Constants
    private const string Ole10Native = "Ole10Native";
    private const string CompObject = "CompObj";
    private const string ObjectInfo = "ObjInfo";
    private const string PrefixOle10Native = "\x0001";
    private const string PrefixCompObject = "\x0001";
    private const string PrefixObjectInfo = "\x0003";
    #endregion
    
    #region Extract
    /// <summary>
    ///     This method saves all the Word embedded binary objects from the <paramref name="inputFile" /> to the
    ///     <paramref name="outputFolder" />
    /// </summary>
    /// <param name="inputFile">The binary Word file</param>
    /// <param name="outputFolder">The output folder</param>
    /// <param name="attachmentsOnly">Sets whether solely attachements shall be extracted or regular OLE elements as well</param>
    /// <param name="continueOnError">Sets whether the extraction should continue when an error, e.g. unsupported AnsiUserType, occurs</param>
    /// <returns></returns>
    /// <exception cref="OEFileIsPasswordProtected">Raised when the <paramref name="inputFile" /> is password protected</exception>
    internal List<string> Extract(string inputFile, string outputFolder, bool attachmentsOnly = false, bool continueOnError = false)
    {
        Logger.WriteToLog("The file is a binary Word document");

        var idOle10Native = (attachmentsOnly ? PrefixOle10Native : "") + Ole10Native;
        var idCompOb = (attachmentsOnly ? PrefixCompObject : "") + CompObject;
        var idObjInfo = (attachmentsOnly ? PrefixObjectInfo : "") + ObjectInfo;

        using var compoundFile = RootStorage.Open(inputFile, FileMode.Open, FileAccess.ReadWrite, StorageModeFlags.Transacted);
        var result = new List<string>();

        if (!compoundFile.TryOpenStorage("ObjectPool", out var objectPool))
            return result;

        Logger.WriteToLog("Object Pool stream found");

        foreach(var item in objectPool.EnumerateEntries())
        {
            try
            {
                if (item.Type != EntryType.Storage) continue;

                if (!objectPool.TryOpenStorage(item.Name, out var childStorage)) continue;

                string extractedFileName;

                if (!childStorage.TryOpenStream(idOle10Native, out _))
                {
                    Logger.WriteToLog("Ole10Native stream found");

                    if (childStorage.TryOpenStream(idCompOb, out var compObj))
                    {
                        Logger.WriteToLog("CompObj stream found");

                        var compObjStream = new CompObjStream(compObj);
                        if (compObjStream.AnsiUserType == "OLE Package")
                        {
                            Logger.WriteToLog("CompObj is of the ansi user type 'OLE Package'");
                            extractedFileName = Extraction.SaveFromStorageNode(childStorage, outputFolder, null);
                            if (!string.IsNullOrEmpty(extractedFileName)) result.Add(extractedFileName!);
                            continue;
                        }

                        Logger.WriteToLog($"CompObj is of the ansi user type '{compObjStream.AnsiUserType}' ... ignoring");
                    }

                    if (childStorage.TryOpenStream(idObjInfo, out var objInfo))
                    {
                        Logger.WriteToLog("ObjInfo stream found");

                        var objInfoStream = new ObjInfoStream(objInfo);
                        // We don't want to export linked objects and objects that are not shown as an icon... 
                        // because these objects are already visible on the Word document
                        if (objInfoStream.Link || !objInfoStream.Icon)
                        {
                            if (objInfoStream.Link)
                                Logger.WriteToLog("ObjInfo stream is a link ... ignoring");

                            if (objInfoStream.Icon)
                                Logger.WriteToLog("ObjInfo stream is an icon ... ignoring");

                            continue;
                        }
                    }

                    extractedFileName = Extraction.SaveFromStorageNode(childStorage, outputFolder, null);
                }
                else
                {
                    Logger.WriteToLog("ObjInfo stream found");

                    // Get the objInfo stream to check if this is a linked file... if so then ignore it
                    if (!childStorage.TryOpenStream(idObjInfo, out var objInfo)) continue;
                    var objInfoStream = new ObjInfoStream(objInfo);

                    // We don't want to export linked objects and objects that are not shown as an icon... 
                    // because these objects are already visible on the Word document
                    if (objInfoStream.Link || !objInfoStream.Icon) continue;
                    extractedFileName = Extraction.SaveFromStorageNode(childStorage, outputFolder);
                }

                if (!string.IsNullOrEmpty(extractedFileName)) result.Add(extractedFileName!);
            }
            catch (Exception ex)
            {
                HandleException(ex, "Word", shallThrow: !continueOnError);
            }
        }

        return result;
    }
    #endregion
}