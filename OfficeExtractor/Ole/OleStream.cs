using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeExtractor.Exceptions;
using OfficeExtractor.Helpers;

namespace OfficeExtractor.Ole
{
    internal class OleStream
    {
        #region Properties
        /// <summary>
        /// This MUST be set to 0x02000001. Otherwise, the OLEStream structure is invalid
        /// </summary>
        public uint Version { get; private set; }

        /// <summary>
        /// This MUST be set to <see cref="OleObjectFormat.Link"/> (0x00000001) or <see cref="OleObjectFormat.File"/> (0x00000002). 
        /// Otherwise, the ObjectHeader structure is invalid
        /// </summary>
        /// <remarks>
        /// 0x00000001 = The ObjectHeader structure MUST be followed by a LinkedObject structure.
        /// 0x00000002 = The ObjectHeader structure MUST be followed by an EmbeddedObject structure.
        /// </remarks>
        public OleObjectFormat Format { get; private set; }

        /// <summary>
        /// This field contains an implementation-specific hint supplied by the application or higher-level 
        /// protocol responsible for creating the data structure. The hint MAY be ignored on processing of 
        /// this data structure
        /// </summary>
        /// <remarks>
        /// If <see cref="Format"/> is set to <see cref="OleObjectFormat.File"/> then this property is empty
        /// </remarks>
        public UInt32 LinkUpdateOptions { get; private set; }
        #endregion

        #region Constructor
        /// <summary>
        /// Creates this object and sets all its properties
        /// </summary>
        /// <param name="binaryReader"></param>
        internal OleStream(BinaryReader binaryReader)
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
                    Format = OleObjectFormat.File;
                    break;

                case 0x00000001:
                case 0x00001001:
                    Format = OleObjectFormat.Link;
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
            if (Format != OleObjectFormat.File)
            {
                // RelativeSourceMonikerStreamSize (4 bytes): This MUST be set to the size, in bytes, of the RelativeSourceMonikerStream field. 
                // If this field has a value 0x00000000, the RelativeSourceMonikerStream field MUST NOT be present.
                var relativeSourceMonikerStreamSize = (int) binaryReader.ReadUInt32();

                // RelativeSourceMonikerStream (variable): This MUST be a MONIKERSTREAM structure that specifies the relative 
                // path to the linked object.

                // AbsoluteSourceMonikerStreamSize (4 bytes): This MUST be set to the size, in bytes, of the AbsoluteSourceMonikerStream field. 
                // This field MUST NOT contain the value 0x00000000.
                var absoluteSourceMonikerStreamSize = (int) binaryReader.ReadUInt32();

                // AbsoluteSourceMonikerStream (variable): This MUST be a MONIKERSTREAM structure that specifies the full path 
                // to the linked object.

                // If the RelativeSourceMonikerStream field is present, it MUST be used by the container application instead of the 
                // AbsoluteSourceMonikerStream. If the RelativeSourceMonikerStream field is not present, the AbsoluteSourceMonikerStream MUST be used 
                // by the container application.

                // ClsidIndicator (4 bytes): This MUST be the LONG as specified in section value -1. Otherwise the OLEStream 
                // structure is invalid.

                // Clsid (16 bytes): This MUST be the CLSID (Packet) containing the object class GUID of the creating application.

                // ReservedDisplayName (4 bytes): This MUST be a LengthPrefixedUnicodeString that can contain any arbitrary value 
                // and MUST be ignored on processing.

                // Reserved2 (4 bytes): This can contain any arbitrary value and MUST be ignored on processing.

                // LocalUpdateTime (4 bytes): This MUST be a FILETIME (Packet) that contains the time when the container application 
                // last updated the RemoteUpdateTime field.

                // LocalCheckUpdateTime (4 bytes): This MUST be a FILETIME (Packet) that contains the time when the container application last 
                // checked the update time of the linked object.

                // RemoteUpdateTime (4 bytes): This MUST be a FILETIME (Packet) that contains the time when the linked object was last updated.
            }
        }
        #endregion
    }
}
