using System.IO;
using System.Text;

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
    ///     The MONIKERSTREAM structure specifies the reference to the linked object.
    /// </summary>
    internal class MonikerStream
    {
        #region Properties
        /// <summary>
        ///     This MUST be the packetized CLSID of an implementation-specific object capable of processing the
        ///     data contained in the StreamData field.
        /// </summary>
        public CLSID Clsid { get; private set; }

        /// <summary>
        ///     This MUST be an array of bytes that specifies the reference to the linked object. The value of
        ///     this array is interpreted in an implementation-specific manner.
        /// </summary>
        public byte[] StreamData { get; private set; }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets all its properties
        /// </summary>
        /// <param name="binaryReader"></param>
        /// <param name="size">The size of the monikerstream</param>
        internal MonikerStream(BinaryReader binaryReader, int size)
        {
            // 16 bytes
            Clsid = new CLSID(binaryReader);
            StreamData = binaryReader.ReadBytes(size - 16);
        }
        #endregion

        #region ToString
        /// <summary>
        ///     Returns the <see cref="StreamData" /> as a string
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Encoding.UTF8.GetString(StreamData);
        }
        #endregion
    }
}