using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
        }
        #endregion
    }
}
