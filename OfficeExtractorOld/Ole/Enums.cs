//
// Enum.cs
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

namespace OfficeExtractor.Ole
{
    #region Enum OleObjectFormat
    /// <summary>
    ///     The format of the embedded object that is stored inside an OLE1 or 
    ///     OLE native container
    /// </summary>
    internal enum OleFormat
    {
        /// <summary>
        ///     The format is NOT set
        /// </summary>
        NotSet = 0x00000000,

        /// <summary>
        ///     The embedded object is a link
        /// </summary>
        Link = 0x00000001,

        /// <summary>
        ///     The embedded object is a file
        /// </summary>
        File = 0x00000002,

        /// <summary>
        ///     The embedded object is a presentation (e.g. an image)
        /// </summary>
        Presentation = 0x00000005
    }
    #endregion

    #region Enum OleClipboardFormat
    /// <summary>
    ///     The OLE version 1.0 and 2.0 clipboard formats
    /// </summary>
    internal enum OleClipboardFormat
    {
        /// <summary>
        ///     The format is a registered clipboard format
        /// </summary>
        Registered = 0x00000000,

        // ReSharper disable InconsistentNaming
        /// <summary>
        ///     Bitmap16 Object structure
        /// </summary>
        CF_BITMAP = 0x00000002,

        /// <summary>
        /// </summary>
        CF_METAFILEPICT = 0x00000003,

        /// <summary>
        ///     DeviceIndependentBitmap Object structure
        /// </summary>
        CF_DIB = 0x00000008,

        /// <summary>
        ///     Enhanced Metafile
        /// </summary>
        CF_ENHMETAFILE = 0x0000000E
        // ReSharper restore InconsistentNaming
    }
    #endregion

    #region Enum OleCF
    /// <summary>
    ///     An unsigned integer that specifies the format this OLE object uses to transmit data to 
    ///     the host application
    /// </summary>
    internal enum OleCf
    {
        UnSpecified = 0x0000,
        RichTextFormat = 0x0001,
        TextFormat = 0x0002,
        MetaOrEnhancedMetaFile = 0x0003,
	    BitMap = 0x0004,
	    DeviceIndependentBitmap = 0x0005,
        HtmlFormat = 0x000A,
	    UnicodeTextFormat = 0x0014
    }
    #endregion
}