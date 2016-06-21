using System;
using System.Collections;
using System.IO;
using CompoundFileStorage;
using CompoundFileStorage.Interfaces;

/*
   Copyright 2014-2016 Kees van Spelde

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
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets all its properties
        /// </summary>
        /// <param name="stream">The Compound File Storage CompObj <see cref="CFStream" /></param>
        internal ObjInfoStream(ICFStream stream)
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
            }
        }
        #endregion
    }
}
