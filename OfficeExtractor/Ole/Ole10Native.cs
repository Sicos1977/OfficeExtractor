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
    /// <param name="skipPaintbrushObjects">
    ///     Sets whether a Paintbrush (PBrush) object shall be skipped. When skipped, <see cref="Format" /> is left
    ///     at <see cref="OleFormat.NotSet" /> and <see cref="NativeData" /> is not read
    /// </param>
    internal Ole10Native(Storage storage, bool skipPaintbrushObjects)
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
                if (skipPaintbrushObjects)
                {
                    // Format stays NotSet, so the object is ignored by the caller
                    Logger.WriteToLog($"Ignoring Ole10Native type '{compObjStream.AnsiUserType}' because skipping Paintbrush objects is requested");
                    break;
                }

                var pbBrushData = ReadNativeData(ole10Native);
                if (pbBrushData == null)
                    break;

                FileName = "Embedded PBrush image.bmp";
                Format = OleFormat.File;
                NativeData = RepairBitmapFileSize(pbBrushData);
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

    #region ReadNativeData
    /// <summary>
    ///     Reads the native data from an <c>\x0001Ole10Native</c> stream that does not contain a package,
    ///     e.g. the one of a Paintbrush object.
    /// </summary>
    /// <remarks>
    ///     The stream starts with a 4-byte little-endian unsigned integer that holds the size of the native
    ///     data that follows it. The stream itself can be longer than that because of padding, so only the
    ///     declared number of bytes is returned. When the declared size is missing or larger than the
    ///     available data, all the data after the size field is returned instead.
    /// </remarks>
    /// <param name="stream">The <c>\x0001Ole10Native</c> stream</param>
    /// <returns>The native data or <c>null</c> when the stream contains no data</returns>
    private static byte[] ReadNativeData(Stream stream)
    {
        const int sizeFieldLength = 4;

        var available = stream.Length - sizeFieldLength;
        if (available <= 0)
            return null;

        stream.Position = 0;
        var sizeField = ReadBytes(stream, sizeFieldLength);
        var declaredSize = BitConverter.ToUInt32(sizeField, 0);

        long size;
        if (declaredSize == 0 || declaredSize > available)
        {
            Logger.WriteToLog($"Ole10Native declared size '{declaredSize}' is invalid, using the available size '{available}' instead");
            size = available;
        }
        else
            size = declaredSize;

        return ReadBytes(stream, (int)size);
    }

    /// <summary>
    ///     Reads exactly <paramref name="count" /> bytes from the current position of the <paramref name="stream" />
    /// </summary>
    /// <param name="stream">The stream to read from</param>
    /// <param name="count">The number of bytes to read</param>
    /// <returns>The read bytes</returns>
    /// <exception cref="EndOfStreamException">Raised when the stream ends before all bytes are read</exception>
    private static byte[] ReadBytes(Stream stream, int count)
    {
        var buffer = new byte[count];
        var offset = 0;

        while (offset < count)
        {
            var read = stream.Read(buffer, offset, count - offset);
            if (read == 0)
                throw new EndOfStreamException($"Expected {count} bytes but the stream ended after {offset} bytes");
            offset += read;
        }

        return buffer;
    }
    #endregion

    #region RepairBitmapFileSize
    /// <summary>
    ///     Makes sure the file size in the BITMAPFILEHEADER of a BMP file matches the real length of the file.
    /// </summary>
    /// <remarks>
    ///     Paintbrush objects store a complete BMP file as native data, but the data can be padded, and so be
    ///     longer than the file size in the header. Strict decoders like ImageMagick reject a BMP file when the
    ///     two values differ ("length and filesize do not match").
    ///     <list type="bullet">
    ///         <item>
    ///             When the data is longer than the header says and the header value still covers the
    ///             pixel data offset, the trailing bytes are removed.
    ///         </item>
    ///         <item>In every other case of a mismatch, the header is set to the length of the data.</item>
    ///     </list>
    ///     Data that is not a BMP file is returned unchanged.
    /// </remarks>
    /// <param name="data">The bitmap data, including the BITMAPFILEHEADER</param>
    /// <returns>The data with a consistent file size</returns>
    private static byte[] RepairBitmapFileSize(byte[] data)
    {
        // BITMAPFILEHEADER: bfType (2), bfSize (4), bfReserved1 (2), bfReserved2 (2), bfOffBits (4)
        const int fileHeaderLength = 14;
        const int sizeOffset = 2;
        const int pixelDataOffset = 10;

        if (data.Length < fileHeaderLength || data[0] != 'B' || data[1] != 'M')
            return data;

        var headerSize = BitConverter.ToUInt32(data, sizeOffset);
        if (headerSize == data.Length)
            return data;

        var pixelOffset = BitConverter.ToUInt32(data, pixelDataOffset);

        if (headerSize < data.Length && headerSize > pixelOffset && pixelOffset >= fileHeaderLength)
        {
            Logger.WriteToLog($"Removing {data.Length - headerSize} padding bytes after the bitmap data");
            var trimmed = new byte[headerSize];
            Array.Copy(data, trimmed, trimmed.Length);
            return trimmed;
        }

        Logger.WriteToLog($"Correcting the bitmap file size in the header from '{headerSize}' to '{data.Length}'");
        var lengthBytes = BitConverter.GetBytes((uint)data.Length);
        Array.Copy(lengthBytes, 0, data, sizeOffset, lengthBytes.Length);
        return data;
    }
    #endregion
}