using System.Collections.Generic;

/*
   Copyright 2014-2015 Kees van Spelde

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
    /// Excel RC4 encryption or decryption
    /// </summary>
    internal class RC4
    {
        #region Fields
        private readonly byte[] _bytes = new byte[256];
        private int _i;
        private int _j;
        #endregion

        #region Constructor
        public RC4(IList<byte> key)
        {
            var keyLength = key.Count;

            for (var i = 0; i < 256; i++)
                _bytes[i] = (byte) i;

            for (int i = 0, j = 0; i < 256; i++)
            {
                j = (j + key[i%keyLength] + _bytes[i]) & 255;
                var temp = _bytes[i];
                _bytes[i] = _bytes[j];
                _bytes[j] = temp;
            }

            _i = 0;
            _j = 0;
        }
        #endregion

        #region Output
        public byte Output()
        {
            _i = (_i + 1) & 255;
            _j = (_j + _bytes[_i]) & 255;

            var temp = _bytes[_i];
            _bytes[_i] = _bytes[_j];
            _bytes[_j] = temp;

            return _bytes[(_bytes[_i] + _bytes[_j]) & 255];
        }
        #endregion

        #region Encrypt
        public void Encrypt(byte[] inputBytes)
        {
            for (var i = 0; i < inputBytes.Length; i++)
            {
                inputBytes[i] = (byte) (inputBytes[i] ^ Output());
            }
        }
        #endregion

        #region Encrypt
        public void Encrypt(byte[] inputBytes, int offSet, int length)
        {
            var end = offSet + length;
            for (var i = offSet; i < end; i++)
                inputBytes[i] = (byte) (inputBytes[i] ^ Output());
        }
        #endregion
    }
}