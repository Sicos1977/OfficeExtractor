using System;
using System.Collections;
using System.IO;
using OpenMcdf;

//
// ObjInfoStream.cs
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
    /// <summary>
    /// Each storage within the ObjectPool storage contains a stream whose name is "\003ObjInfo" where \003 is 
    /// the character with value 0x0003, not the string literal "\003". This stream contains an ODT structure 
    /// which specifies information about that embedded OLE object.
    /// </summary>
    internal class ObjInfoStream
    {
        #region Properties
        /// <summary>
        /// If this is true, then the application MUST assume that this OLE object’s class identifier 
        /// (CLSID) is {00020907-0000-0000-C000-000000000046}.
        /// </summary>
        internal bool DefHandler { get; private set; }

        /// <summary>
        /// Specifies whether this OLE object is a link.
        /// </summary>
        internal bool Link { get; private set; }

        /// <summary>
        /// Specifies whether this OLE object is being represented by an icon
        /// </summary>
        internal bool Icon { get; private set; }

        /// <summary>
        /// Specifies whether this OLE object is only compatible with OLE 1. 
        /// If this is false, then the object is compatible with OLE 2.
        /// </summary>
        internal bool IsOle1 { get; private set; }

        /// <summary>
        /// Specifies whether the user has requested that this OLE object only be updated in response to 
        /// a user action. If <see cref="Manual"/> is <c>false</c>, then the user has requested that 
        /// this OLE object update automatically. If <see cref="Link"/> is <c>false</c>, then <see cref="Manual"/> 
        /// is undefined and MUST be ignored.
        /// </summary>
        internal bool Manual { get; private set; }

        /// <summary>
        /// Specifies whether this OLE object has requested to be notified when it is resized by its container.
        /// </summary>
        internal bool RecomposeOnResize { get; private set; }

        /// <summary>
        /// Specifies whether this object is an OLE control
        /// </summary>
        internal bool Ocx { get; private set; }

        /// <summary>
        /// If <see cref="Ocx"/> is <c>false</c>, then this MUST be <c>false</c>. If <see cref="Ocx"/> is <c>true</c>, 
        /// then <see cref="Stream"/> is a boolean that specifies whether this OLE control stores its data in a single 
        /// stream instead of a storage. If <see cref="Stream"/> is <c>true</c>, then the data for the OLE control is 
        /// in a stream called "\003OCXDATA" where \003 is the character with value 0x0003, not the string literal "\003"
        /// </summary>
        internal bool Stream { get; private set; }

        /// <summary>
        /// Specifies whether this OLE object supports the IViewObject interface.
        /// </summary>
        internal bool ViewObject { get; private set; }

        /// <summary>
        /// An unsigned integer that specifies the format this OLE object uses to transmit data to the host application
        /// </summary>
        internal OleCf Cf { get; private set; }

        /// <summary>
        /// A bit that specifies that the presentation of this OLE object in the document is in the Enhanced Metafile format. 
        /// This is different from fStoredAsEMF in the case of an object being represented as an icon. For icons, the icon can 
        /// be an Enhanced Metafile even if the OLE object does not support the Enhanced Metafile format.
        /// </summary>
        internal bool Emf { get; private set; }

        /// <summary>
        /// A bit that specifies whether the application that saved this Word Binary file had queried this OLE object to determine 
        /// whether it supported the Enhanced Metafile format.
        /// </summary>
        internal bool QueriedEmf { get; private set; }

        /// <summary>
        /// A bit that specifies that this OLE object supports the Enhanced Metafile format.
        /// </summary>
        internal bool StoredAsEmf { get; private set; }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets all its properties
        /// </summary>
        /// <param name="stream">The Compound File Storage CompObj <see cref="CFStream" /></param>
        internal ObjInfoStream(CFStream stream)
        {
            using (var memoryStream = new MemoryStream(stream.GetData()))
            using (var binaryReader = new BinaryReader(memoryStream))
            {
                var bytes = binaryReader.ReadBytes(2);
                var bitArray = new BitArray(bytes);

                // A - reserved1 (1 bit): Undefined and MUST be ignored.

                // B - fDefHandler (1 bit): If this bit is 1, then the application MUST assume that this OLE object’s 
                //     class identifier (CLSID) is {00020907-0000-0000-C000-000000000046}.
                DefHandler = bitArray.Get(1);

                // C - reserved2 (1 bit): Undefined and MUST be ignored.
                // D - reserved3 (1 bit): Undefined and MUST be ignored.

                // E - fLink (1 bit): A bit that specifies whether this OLE object is a link.
                Link = bitArray.Get(4);

                // F - reserved4 (1 bit): Undefined and MUST be ignored.
                // G - fIcon (1 bit): A bit that specifies whether this OLE object is being represented by an icon.
                Icon = bitArray.Get(6);

                // H - fIsOle1 (1 bit): A bit that specifies whether this OLE object is only compatible with OLE 1. 
                //     If this bit is zero, then the object is compatible with OLE 2.
                IsOle1 = bitArray.Get(7);

                // I - fManual (1 bit): A bit that specifies whether the user has requested that this OLE object only 
                //     be updated in response to a user action. If fManual is zero, then the user has requested that 
                //     this OLE object update automatically. If fLink is zero, then fManual is undefined and MUST be ignored.
                Manual = bitArray.Get(8);

                // J - fRecomposeOnResize (1 bit): A bit that specifies whether this OLE object has requested to be notified 
                //     when it is resized by its container.
                RecomposeOnResize = bitArray.Get(9);

                // K - reserved5 (1 bit): MUST be zero and MUST be ignored.
                // L - reserved6 (1 bit): MUST be zero and MUST be ignored.
                // M - fOCX (1 bit): A bit that specifies whether this object is an OLE control.
                Ocx = bitArray.Get(12);

                // N - fStream (1 bit): If fOCX is zero, then this bit MUST be zero. If fOCX is 1, then fStream is a bit that 
                //     specifies whether this OLE control stores its data in a single stream instead of a storage. If fStream 
                //     is 1, then the data for the OLE control is in a stream called "\003OCXDATA" where \003 is the character 
                //     with value 0x0003, not the string literal "\003".
                Stream = bitArray.Get(13);

                // O - reserved7 (1 bit): Undefined and MUST be ignored.
                // P - fViewObject (1 bit): A bit that specifies whether this OLE object supports the IViewObject interface.
                ViewObject = bitArray.Get(15);

                try
                {
                    Cf = (OleCf) binaryReader.ReadUInt16();
                }
                catch (Exception)
                {
                    Cf = OleCf.UnSpecified;
                }

                try
                {
                    bytes = binaryReader.ReadBytes(2);
                    bitArray = new BitArray(bytes);

                    // A - fEMF(1 bit): A bit that specifies that the presentation of this OLE object in the document is in the 
                    //     Enhanced Metafile format. This is different from fStoredAsEMF in the case of an object being represented 
                    //     as an icon.For icons, the icon can be an Enhanced Metafile even if the OLE object does not support the 
                    //     Enhanced Metafile format.
                    Emf =  bitArray.Get(1);

                    // B - reserved1(1 bit): MUST be zero and MUST be ignored.

                    // C - fQueriedEMF(1 bit): A bit that specifies whether the application that saved this Word Binary file had 
                    //     queried this OLE object to determine whether it supported the Enhanced Metafile format.
                    QueriedEmf = bitArray.Get(3);

                    // D - fStoredAsEMF(1 bit): A bit that specifies that this OLE object supports the Enhanced Metafile format.
                    StoredAsEmf = bitArray.Get(4);

                    // E - reserved2(1 bit): Undefined and MUST be ignored.
                    // F - reserved3(1 bit): Undefined and MUST be ignored.
                    // reserved4(10 bits): Undefined and MUST be ignored.
                }
                catch (Exception)
                {
                    // Ignore
                }
            }
        }
        #endregion
    }
}
