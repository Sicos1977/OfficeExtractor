using System;
using System.IO;
using System.Linq;
using CompoundFileStorage.Interfaces;
using OfficeExtractor.Helpers;

/*
   Copyright 2013 - 2016 Kees van Spelde

   Licensed under The Code Project Open License (CPOL) 1.02;
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.codeproject.com/info/cpol10.aspx

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

namespace OfficeExtractor.Ole
{
    /// <summary>
    ///     The AttachDesc stream stores information about the attachment. The following table specifies
    ///     the format of the fields of the AttachDesc stream in the order in which they appear. Some of
    ///     the fields contain values of Attachment object properties.
    /// </summary>
    internal class AttachDescStream
    {
        #region Properties
        /// <summary>
        ///     Contains the version. When creating a rights-managed email message, this value MUST
        ///     always be set to 0x0203. The value is stored in the little-endian format.
        /// </summary>
        public uint Version { get; private set; }

        /// <summary>
        ///     SHOULD contain the value of the PidTagAttachLongPathname property of the
        ///     attachment, if present; otherwise, it MUST be 0x00.
        /// </summary>
        public string LongPathName { get; private set; }

        /// <summary>
        ///     MUST contain the value of the PidTagAttachPathname property sof the attachment,
        ///     if present; otherwise, it MUST be 0x00.
        /// </summary>
        public string PathName { get; private set; }

        /// <summary>
        ///     MUST contain the value of the PidTagDisplayName property of the attachment,
        ///     if present; otherwise, it MUST be 0x00.
        /// </summary>
        public string DisplayName { get; private set; }

        /// <summary>
        ///     MUST contain the value of the PidTagAttachLongFilename property of the attachment,
        ///     if present; otherwise, it MUST be 0x00.
        /// </summary>
        public string LongFileName { get; private set; }

        /// <summary>
        ///     MUST contain the value of the PidTagAttachFilename property of the attachment,
        ///     if present; otherwise, it MUST be 0x00.
        /// </summary>
        public string FileName { get; private set; }

        /// <summary>
        ///     MUST contain the value of the PidTagAttachExtension property of the attachment,
        ///     if present; otherwise, it MUST be 0x00.
        /// </summary>
        public string Extension { get; private set; }

        /// <summary>
        ///     MUST contain the value of the PidTagCreationTime property of the attachment,
        ///     if present; otherwise, it MUST be 0x0000000000000000. This is stored in little-endian format.
        /// </summary>
        public DateTime FileCreationTime { get; private set; }

        /// <summary>
        ///     MUST contain the value of the PidTagLastModificationTime property of the attachment, if present;
        ///     otherwise, it MUST be 0x0000000000000000. This is stored in little-endian format.
        /// </summary>
        public DateTime FileLastModifiedTime { get; private set; }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets all its properties
        /// </summary>
        /// <param name="stream">The Compound File Storage AttachDesc <see cref="CFStream" /></param>
        internal AttachDescStream(ICFStream stream)
        {
            using (var memoryStream = new MemoryStream(stream.GetData()))
            using (var binaryReader = new BinaryReader(memoryStream))
            {
                Version = binaryReader.ReadUInt16();
                LongPathName = Strings.Read1ByteLengthPrefixedAnsiString(binaryReader);
                PathName = Strings.Read1ByteLengthPrefixedAnsiString(binaryReader);
                DisplayName = Strings.Read1ByteLengthPrefixedAnsiString(binaryReader);
                LongFileName = Strings.Read1ByteLengthPrefixedAnsiString(binaryReader);
                FileName = Strings.Read1ByteLengthPrefixedAnsiString(binaryReader);
                Extension = Strings.Read1ByteLengthPrefixedAnsiString(binaryReader);
                var fileCreationTime = binaryReader.ReadBytes(8).ToArray();
                FileCreationTime = DateTime.FromFileTime(BitConverter.ToInt64(fileCreationTime, 0));
                var fileLastModifiedTime = binaryReader.ReadBytes(8).ToArray();
                FileLastModifiedTime = DateTime.FromFileTime(BitConverter.ToInt64(fileLastModifiedTime, 0));
            }
        }
        #endregion
    }
}