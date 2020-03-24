using System;
using System.IO;
using System.Linq;
using OpenMcdf;

//
// OleStream.cs
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
    ///     The OLEStream structure is contained inside an OLE Compound File Stream object. The name of this
    ///     Compound File Stream object is "\1Ole". The stream object is contained within the OLE Compound File
    ///     Storage object corresponding to the linked object or embedded object. The OLEStream structure specifies
    ///     whether the storage object is for a linked object or an embedded object. When this structure specifies a
    ///     storage object for a linked object, it also specifies the reference to the linked object.
    /// </summary>
    internal class OleStream
    {
        #region Properties
        /// <summary>
        ///     This MUST be set to 0x02000001. Otherwise, the OLEStream structure is invalid
        /// </summary>
        public uint Version { get; private set; }

        /// <summary>
        ///     This MUST be set to <see cref="OleFormat.Link" /> (0x00000001) or <see cref="OleFormat.File" />
        ///     (0x00000002).
        ///     Otherwise, the ObjectHeader structure is invalid
        /// </summary>
        /// <remarks>
        ///     0x00000001 = The ObjectHeader structure MUST be followed by a LinkedObject structure.
        ///     0x00000002 = The ObjectHeader structure MUST be followed by an EmbeddedObject structure.
        /// </remarks>
        public OleFormat Format { get; private set; }

        /// <summary>
        ///     This field contains an implementation-specific hint supplied by the application or higher-level
        ///     protocol responsible for creating the data structure. The hint MAY be ignored on processing of
        ///     this data structure
        /// </summary>
        /// <remarks>
        ///     Only available when the <see cref="Format" /> is set to <see cref="OleFormat.Link" />, otherwise <c>null</c>
        /// </remarks>
        public UInt32 LinkUpdateOptions { get; private set; }

        /// <summary>
        ///     RelativeSourceMonikerStream (variable): This MUST be a MONIKERSTREAM structure that specifies the relative
        ///     path to the linked object.
        /// </summary>
        /// <remarks>
        ///     Only available when the <see cref="Format" /> is set to <see cref="OleFormat.Link" />, otherwise <c>null</c>
        /// </remarks>
        public MonikerStream RelativeSource { get; private set; }

        /// <summary>
        ///     AbsoluteSourceMonikerStream (variable): This MUST be a MONIKERSTREAM structure that specifies the full path
        ///     to the linked object.
        /// </summary>
        /// <remarks>
        ///     Only available when the <see cref="Format" /> is set to <see cref="OleFormat.Link" />, otherwise <c>null</c>
        /// </remarks>
        public MonikerStream AbsoluteSource { get; private set; }

        /// <summary>
        ///     This MUST be the CLSID (Packet) containing the object class GUID of the creating application.
        /// </summary>
        /// <remarks>
        ///     Only available when the <see cref="Format" /> is set to <see cref="OleFormat.Link" />, otherwise <c>null</c>
        /// </remarks>
        public CLSID Clsid { get; private set; }

        /// <summary>
        ///     LocalUpdateTime (4 bytes): This MUST be a FILETIME (Packet) that contains the time when the container application
        ///     last updated the RemoteUpdateTime field.
        /// </summary>
        public DateTime LocalUpdateTime { get; private set; }

        /// <summary>
        ///     LocalCheckUpdateTime (4 bytes): This MUST be a FILETIME (Packet) that contains the time when the container
        ///     application last
        ///     checked the update time of the linked object.
        /// </summary>
        public DateTime LocalCheckUpdateTime { get; private set; }

        /// <summary>
        ///     RemoteUpdateTime (4 bytes): This MUST be a FILETIME (Packet) that contains the time when the linked object was last
        ///     updated.
        /// </summary>
        public DateTime RemoteUpdateTime { get; private set; }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets all its properties
        /// </summary>
        /// <param name="stream">The Compound File Storage Ole <see cref="CFStream" /></param>
        internal OleStream(CFStream stream)
        {
            using (var memoryStream = new MemoryStream(stream.GetData()))
            using (var binaryReader = new BinaryReader(memoryStream))
            {
                Version = binaryReader.ReadUInt16();

                // Flags (4 bytes): If this field is set to 0x00000001, the OLEStream structure MUST be for a linked object and
                // the CLSID field of the Compound File Directory Entry of the OLE Compound File Storage object MUST be set to 
                // CLSID_StdOleLink ({00000300-0000-0000-C000-000000000046}). If this field is set to 0x00000000, then the OLEStream 
                // structure MUST be for an embedded object and the CLSID field of the Compound File Directory Entry
                // of the OLE Compound File Storage object MUST be set to the object class GUID of the creating application.
                var flags = binaryReader.ReadUInt32();

                switch (flags)
                {
                    case 0x00000000:
                    case 0x00001000:
                        Format = OleFormat.File;
                        break;

                    case 0x00000001:
                    case 0x00001001:
                        Format = OleFormat.Link;
                        break;
                }

                // LinkUpdateOption (4 bytes): This field contains an implementation-specific hint supplied by the application or by
                // a higher-level protocol that creates the data structure. The hint MAY be ignored on processing of this data structure
                LinkUpdateOptions = binaryReader.ReadUInt32();

                //Reserved1 (4 bytes): This MUST be set to 0x00000000. Otherwise, the OLEStream structure is invalid
                binaryReader.ReadUInt32();

                // ReservedMonikerStreamSize (4 bytes): This MUST be set to the size, in bytes, of the ReservedMonikerStream field. If this 
                // field has a value 0x00000000, the ReservedMonikerStream field MUST NOT be present.
                var reservedMonikerStreamSize = (int) binaryReader.ReadUInt32();

                // ReservedMonikerStream (variable): This MUST be a MONIKERSTREAM structure that can contain any arbitrary 
                // value and MUST be ignored on processing.
                binaryReader.ReadBytes(reservedMonikerStreamSize);

                // Note The fields that follow MUST NOT be present if the OLEStream structure is for an embedded object.
                if (Format == OleFormat.Link)
                {
                    // RelativeSourceMonikerStreamSize (4 bytes): This MUST be set to the size, in bytes, of the RelativeSourceMonikerStream field. 
                    // If this field has a value 0x00000000, the RelativeSourceMonikerStream field MUST NOT be present.
                    var relativeSourceMonikerStreamSize = (int) binaryReader.ReadUInt32();

                    // RelativeSourceMonikerStream (variable): This MUST be a MONIKERSTREAM structure that specifies the relative 
                    // path to the linked object.
                    if (relativeSourceMonikerStreamSize > 0)
                        RelativeSource = new MonikerStream(binaryReader, relativeSourceMonikerStreamSize);

                    // AbsoluteSourceMonikerStreamSize (4 bytes): This MUST be set to the size, in bytes, of the AbsoluteSourceMonikerStream field. 
                    // This field MUST NOT contain the value 0x00000000.
                    var absoluteSourceMonikerStreamSize = (int) binaryReader.ReadUInt32();

                    // AbsoluteSourceMonikerStream (variable): This MUST be a MONIKERSTREAM structure that specifies the full path 
                    // to the linked object.
                    if (absoluteSourceMonikerStreamSize > 0)
                        AbsoluteSource = new MonikerStream(binaryReader, absoluteSourceMonikerStreamSize);

                    // If the RelativeSourceMonikerStream field is present, it MUST be used by the container application instead of the 
                    // AbsoluteSourceMonikerStream. If the RelativeSourceMonikerStream field is not present, the AbsoluteSourceMonikerStream MUST be used 
                    // by the container application.

                    // ClsidIndicator (4 bytes): This MUST be the LONG as specified in section value -1. Otherwise the OLEStream 
                    // structure is invalid.
                    binaryReader.ReadUInt32();

                    // Clsid (16 bytes): This MUST be the CLSID (Packet) containing the object class GUID of the creating application.
                    Clsid = new CLSID(binaryReader);

                    // ReservedDisplayName (4 bytes): This MUST be a LengthPrefixedUnicodeString that can contain any arbitrary value 
                    // and MUST be ignored on processing.
                    binaryReader.ReadUInt32();

                    // Reserved2 (4 bytes): This can contain any arbitrary value and MUST be ignored on processing.
                    binaryReader.ReadUInt32();

                    // LocalUpdateTime (4 bytes): This MUST be a FILETIME (Packet) that contains the time when the container application 
                    // last updated the RemoteUpdateTime field.
                    var localUpdateTime = binaryReader.ReadBytes(4).Reverse().ToArray();
                    LocalUpdateTime = DateTime.FromFileTime(BitConverter.ToInt32(localUpdateTime, 0));

                    // LocalCheckUpdateTime (4 bytes): This MUST be a FILETIME (Packet) that contains the time when the container application last 
                    // checked the update time of the linked object.
                    var localCheckUpdateTime = binaryReader.ReadBytes(4).Reverse().ToArray();
                    LocalCheckUpdateTime = DateTime.FromFileTime(BitConverter.ToInt32(localCheckUpdateTime, 0));

                    // RemoteUpdateTime (4 bytes): This MUST be a FILETIME (Packet) that contains the time when the linked object was last updated.
                    var remoteUpdateTime = binaryReader.ReadBytes(4).Reverse().ToArray();
                    RemoteUpdateTime = DateTime.FromFileTime(BitConverter.ToInt32(remoteUpdateTime, 0));
                }
            }
        }
        #endregion
    }
}