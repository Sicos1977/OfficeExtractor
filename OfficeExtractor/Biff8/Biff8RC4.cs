using System;

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
    ///     Used for both encrypting and decrypting BIFF8 streams. The internal <see cref="RC4" />
    ///     instance is renewed (re-keyed) every 1024 bytes.
    /// </summary>
    internal class Biff8RC4
    {
        #region Fields
        private const int Rc4RekeyingInterval = 1024;
        private readonly Biff8EncryptionKey _key;
        private int _currentKeyIndex;
        private int _nextRc4BlockStart;
        private RC4 _rc4;
        private bool _shouldSkipEncryptionOnCurrentRecord;
        private int _streamPos;
        #endregion

        #region Constructor
        public Biff8RC4(int initialOffset, Biff8EncryptionKey key)
        {
            if (initialOffset >= Rc4RekeyingInterval)
            {
                throw new ArgumentException("InitialOffset (" + initialOffset + ")>"
                                            + Rc4RekeyingInterval + " not supported yet");
            }

            _key = key;
            _streamPos = 0;
            RekeyForNextBlock();
            _streamPos = initialOffset;
            for (var i = initialOffset; i > 0; i--)
                _rc4.Output();

            _shouldSkipEncryptionOnCurrentRecord = false;
        }
        #endregion

        #region RekeyForNextBlock
        private void RekeyForNextBlock()
        {
            _currentKeyIndex = _streamPos/Rc4RekeyingInterval;
            _rc4 = _key.CreateRC4(_currentKeyIndex);
            _nextRc4BlockStart = (_currentKeyIndex + 1)*Rc4RekeyingInterval;
        }
        #endregion

        #region GetNextRC4Byte
        private int GetNextRC4Byte()
        {
            if (_streamPos >= _nextRc4BlockStart)
                RekeyForNextBlock();

            var mask = _rc4.Output();
            _streamPos++;
            if (_shouldSkipEncryptionOnCurrentRecord)
                return 0;

            return mask & 0xFF;
        }
        #endregion

        #region StartRecord
        public void StartRecord(int currentSid)
        {
            _shouldSkipEncryptionOnCurrentRecord = IsNeverEncryptedRecord(currentSid);
        }

        private static bool IsNeverEncryptedRecord(int sid)
        {
            switch (sid)
            {
                case 0x809:
                case 0xe1:
                case 0x2F:
                    return true;

                default:
                    return false;
            }
        }
        #endregion

        #region SkipTwoBytes
        public void SkipTwoBytes()
        {
            GetNextRC4Byte();
            GetNextRC4Byte();
        }
        #endregion

        #region Xor
        public void Xor(byte[] bufferBytes, int offSet, int length)
        {
            var nLeftInBlock = _nextRc4BlockStart - _streamPos;
            if (length <= nLeftInBlock)
            {
                // simple case - this read does not cross key blocks
                _rc4.Encrypt(bufferBytes, offSet, length);
                _streamPos += length;
                return;
            }

            var offset = offSet;
            var len = length;

            // Start by using the rest of the current block
            if (len > nLeftInBlock)
            {
                if (nLeftInBlock > 0)
                {
                    _rc4.Encrypt(bufferBytes, offset, nLeftInBlock);
                    _streamPos += nLeftInBlock;
                    offset += nLeftInBlock;
                    len -= nLeftInBlock;
                }
                RekeyForNextBlock();
            }

            // All full blocks following
            while (len > Rc4RekeyingInterval)
            {
                _rc4.Encrypt(bufferBytes, offset, Rc4RekeyingInterval);
                _streamPos += Rc4RekeyingInterval;
                offset += Rc4RekeyingInterval;
                len -= Rc4RekeyingInterval;
                RekeyForNextBlock();
            }

            // Finish with incomplete block
            _rc4.Encrypt(bufferBytes, offset, len);
            _streamPos += len;
        }
        #endregion

        #region XorByte
        public int XorByte(int rawVal)
        {
            var mask = GetNextRC4Byte();
            return (byte) (rawVal ^ mask);
        }
        #endregion

        #region Xorshort
        public int Xorshort(int rawVal)
        {
            var byte0 = GetNextRC4Byte();
            var byte1 = GetNextRC4Byte();
            var mask = (byte1 << 8) + (byte0 << 0);
            return rawVal ^ mask;
        }
        #endregion

        #region XorInt
        public int XorInt(int rawVal)
        {
            var byte0 = GetNextRC4Byte();
            var byte1 = GetNextRC4Byte();
            var byte2 = GetNextRC4Byte();
            var byte3 = GetNextRC4Byte();
            var mask = (byte3 << 24) + (byte2 << 16) + (byte1 << 8) + (byte0 << 0);
            return rawVal ^ mask;
        }
        #endregion

        #region XorLong
        public long XorLong(long rawVal)
        {
            var byte0 = GetNextRC4Byte();
            var byte1 = GetNextRC4Byte();
            var byte2 = GetNextRC4Byte();
            var byte3 = GetNextRC4Byte();
            var byte4 = GetNextRC4Byte();
            var byte5 = GetNextRC4Byte();
            var byte6 = GetNextRC4Byte();
            var byte7 = GetNextRC4Byte();
            var mask =
                (((long) byte7) << 56)
                + (((long) byte6) << 48)
                + (((long) byte5) << 40)
                + (((long) byte4) << 32)
                + (((long) byte3) << 24)
                + (byte2 << 16)
                + (byte1 << 8)
                + (byte0 << 0);
            return rawVal ^ mask;
        }
        #endregion
    }
}