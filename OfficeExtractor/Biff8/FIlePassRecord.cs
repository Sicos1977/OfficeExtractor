using System;
using System.IO;
using OfficeExtractor.Biff8.Interfaces;
using OfficeExtractor.Exceptions;

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

namespace OfficeExtractor.Biff8
{

    /// <summary>
    /// Used to read the FilePass record
    /// </summary>
    internal class FilePassRecord
    {
        #region Consts
        private const short SidId = 0x2F;
        private const int EncryptionXor = 0;
        private const int EncryptionOther = 1;
        private const int EncryptionOtherRC4 = 1;
        private const int EncryptionOtherCapi2 = 2;
        private const int EncryptionOtherCapi3 = 3;
        #endregion

        #region Properties
        public byte[] DocId { get; private set; }

        public byte[] SaltData { get; private set; }

        public byte[] SaltHash { get; private set; }

        public short Sid
        {
            get { return SidId; }
        }
        #endregion

        #region Constructor
        public FilePassRecord(Stream inputStream)
        {
            // ReSharper disable once SuspiciousTypeConversion.Global
            var input = new LittleEndianInputStream(inputStream);
            if (input == null)
                throw new ArgumentNullException("inputStream");

            var encryptionType = input.ReadUShort();
            switch (encryptionType)
            {
                case EncryptionXor:
                    throw new OEExcelConfiguration("XOR obfuscation is not supported");

                case EncryptionOther:
                    break;
                
                default:
                    throw new OEExcelConfiguration("Unknown encryption type " + encryptionType);
            }

            var encryptionInfo = input.ReadUShort();
            switch (encryptionInfo)
            {
                case EncryptionOtherRC4:
                    // handled below
                    break;

                case EncryptionOtherCapi2:
                case EncryptionOtherCapi3:
                    throw new OEExcelConfiguration("CryptoAPI encryption is not supported");
                
                default:
                    throw new OEExcelConfiguration("Unknown encryption info " + encryptionInfo);
            }

            input.ReadUShort();
            DocId = Read(input, 16);
            SaltData = Read(input, 16);
            SaltHash = Read(input, 16);
        }
        #endregion

        #region Read
        /// <summary>
        /// Returns <see cref="size"/> bytes from the <see cref="input"/> stream
        /// </summary>
        /// <param name="input"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        private static byte[] Read(ILittleEndianInput input, int size)
        {
            var result = new byte[size];
            input.ReadFully(result);
            return result;
        }
        #endregion
    }
}