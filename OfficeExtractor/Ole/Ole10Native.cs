using System;
using System.IO;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;
using OpenMcdf;

// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable VariableLengthStringHexEscapeSequence
// ReSharper disable GrammarMistakeInComment
// ReSharper disable CommentTypo

//
// Ole10Native.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2025 Magic-Sessions. (www.magic-sessions.com)
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

namespace OfficeExtractor.Ole;

/// <summary>
///     This class represents an OLE version 2.0 object
/// </summary>
/// <remarks>
///     See the Microsoft documentation at https://msdn.microsoft.com/en-us/library/dd942280.aspx
/// </remarks>
internal class Ole10Native
{
    #region Properties
    /// <summary>
    ///     This MUST be set to <see cref="OleFormat.Link" /> (0x00000001) or <see cref="OleFormat.File" />
    ///     (0x00000002).
    ///     Otherwise, the ObjectHeader structure is invalid
    /// </summary>
    public OleFormat Format { get; private set; }

    /// <summary>
    ///     This MUST be a LengthPrefixedAnsiString which contain a registered clipboard format name
    /// </summary>
    public string StringFormat { get; private set; }

    /// <summary>
    ///     This MUST be a LengthPrefixedAnsiString structure that contains a display name of the linked
    ///     object or embedded object.
    /// </summary>
    public string AnsiUserType { get; private set; }

    /// <summary>
    ///     AnsiClipboardFormat (variable): This MUST be a ClipboardFormatOrAnsiString structure that contains the
    ///     Clipboard Format of the linked object or embedded object.
    /// </summary>
    public OleClipboardFormat ClipboardFormat { get; private set; }

    /// <summary>
    ///     The filename
    /// </summary>
    public string FileName { get; private set; }

    /// <summary>
    ///     The path to the file before it was embedded
    /// </summary>
    public string FilePath { get; private set; }

    /// <summary>
    ///     The content of the embedded file
    /// </summary>
    public byte[] NativeData { get; private set; }
    #endregion

    #region Constructor
    /// <summary>
    ///     Creates this object and sets all its properties
    /// </summary>
    /// <param name="storage">The OLE version 2.0 object as a <see cref="Storage" /></param>
    internal Ole10Native(Storage storage)
    {
        if (storage == null)
            throw new ArgumentNullException(nameof(storage));

        var ole10Native = storage.OpenStream("\x0001Ole10Native");

        // Check if CompObj stream exists
        if (!storage.TryOpenStream("\x0001CompObj", out var compObj))
        {
            Logger.WriteToLog("CompObj stream not found, attempting to parse Ole10Native as Package");
            
            // When CompObj is missing, try to parse as a Package directly
            try
            {
                var package = new Package(ole10Native, 4);
                Format = package.Format;
                FileName = Path.GetFileName(package.FileName);
                FilePath = package.FilePath;
                NativeData = package.Data;
                AnsiUserType = "OLE Package";
            }
            catch (Exception ex)
            {
                Logger.WriteToLog($"Failed to parse Ole10Native without CompObj: {ex.Message}");
                throw new OEObjectTypeNotSupported("Unable to parse Ole10Native stream without CompObj stream", ex);
            }
            return;
        }

        var compObjStream = new CompObjStream(compObj);

        AnsiUserType = compObjStream.AnsiUserType;
        StringFormat = compObjStream.StringFormat;
        ClipboardFormat = compObjStream.ClipboardFormat;

        switch (compObjStream.AnsiUserType)
        {
            case "OLE Package":
                var package = new Package(ole10Native, 4);
                Format = package.Format;
                FileName = Path.GetFileName(package.FileName);
                FilePath = package.FilePath;
                NativeData = package.Data;
                break;

            case "PBrush":
            case "Paintbrush-Bild":
            case "Paintbrush-afbeelding":
                var pbBrushSize = (int)ole10Native.Length - 4;
                if (pbBrushSize <= 0)
                    break;
                var pbBrushData = new byte[pbBrushSize];
                ole10Native.Position = 4;
                _ = ole10Native.Read(pbBrushData, 0, pbBrushSize);
                FileName = "Embedded PBrush image.bmp";
                Format = OleFormat.File;
                NativeData = pbBrushData;
                break;

            case "Pakket":
                Logger.WriteToLog("Ignoring Ole10Native type 'Pakket'");
                break;

            // MathType (http://docs.wiris.com/en/mathtype/start) is a equations editor
            // The data is stored in the MTEF format within image file formats (PICT, WMF, EPS, GIF) or Office documents
            // as kind of pickaback data. (http://docs.wiris.com/en/mathtype/mathtype_desktop/mathtype-sdk/mtefstorage).
            // Within Office, a placeholder image shows the created equation.
            // Because MathType does not support storing equations in a separate MTEF file, a export of the data is not
            // directly possible and would require a conversion into the mentioned file formats.
            // Due that facts, it make no sense try to export the data.
            case "MathType 5.0 Equation":
                Logger.WriteToLog("Ignoring Ole10Native type 'MathType 5.0 Equation'");
                break;

            // Used by the depreciated Microsoft Office ClipArt Gallery
            // supposedly to store some metadata
            case "MS_ClipArt_Gallery":
                Logger.WriteToLog("MS_ClipArt_Gallery'");
                break;

            case "Microsoft ClipArt Gallery":
                Logger.WriteToLog("Ignoring Ole10Native type 'Microsoft ClipArt Gallery'");
                break;

            case "Bitmap Image":
                Logger.WriteToLog("Ignoring Ole10Native type 'Bitmap Image'");
                break;

            default:
                throw new OEObjectTypeNotSupported($"Unsupported OleNative AnsiUserType '{compObjStream.AnsiUserType}' found");
        }
    }
    #endregion
}