using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using OfficeExtractor.Exceptions;

namespace OfficeExtractor.Biff8
{
    /// <summary>
    /// Used to create or validate the Excel encryption key
    /// </summary>
    internal class Biff8EncryptionKey
    {
        #region Consts
        // These two constants coincidentally have the same value
        private const int KeyDigestLength = 5;
        private const int PasswordHashNumberOfBytesUsed = 5;
        #endregion

        #region Fields
        [ThreadStatic]
        private static String _userPasswordTls;
        private readonly byte[] _keyDigest;
        #endregion

        #region Properties
        /// <summary>
        /// Returns the BIFF8 encryption/decryption password for the current thread, <code>null</code> if it is currently unSet.
        /// </summary>
        public static String CurrentUserPassword
        {
            get { return _userPasswordTls; }
            set { _userPasswordTls = value; }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Create using the default password and a specified docId
        /// </summary>
        /// <param name="docId"></param>
        /// <returns></returns>
        public static Biff8EncryptionKey Create(byte[] docId)
        {
            return new Biff8EncryptionKey(CreateKeyDigest("VelvetSweatshop", docId));
        }

        public static Biff8EncryptionKey Create(string password, byte[] docIdData)
        {
            return new Biff8EncryptionKey(CreateKeyDigest(password, docIdData));
        }
        #endregion

        #region Biff8EncryptionKey
        private Biff8EncryptionKey(byte[] keyDigest)
        {
            if (keyDigest.Length != KeyDigestLength)
                throw new OEFileIsCorrupt("Expected 5 byte key digest, but got " + keyDigest.Length);

            _keyDigest = keyDigest;
        }
        #endregion

        #region CreateKeyDigest
        private static byte[] CreateKeyDigest(String password, byte[] docIdData)
        {
            Check16Bytes(docIdData, "docId");
            var nChars = Math.Min(password.Length, 16);
            var passwordData = new byte[nChars*2];
            for (var i = 0; i < nChars; i++)
            {
                var chr = password[i];
                passwordData[i*2 + 0] = (byte) ((chr << 0) & 0xFF);
                passwordData[i*2 + 1] = (byte) ((chr << 8) & 0xFF);
            }

            using (MD5 md5 = new MD5CryptoServiceProvider())
            {
                var passwordHash = md5.ComputeHash(passwordData);

                md5.Initialize();

                var data = new byte[PasswordHashNumberOfBytesUsed*16 + docIdData.Length*16];

                var offset = 0;
                for (var i = 0; i < 16; i++)
                {
                    Array.Copy(passwordHash, 0, data, offset, PasswordHashNumberOfBytesUsed);
                    offset += PasswordHashNumberOfBytesUsed;
                    Array.Copy(docIdData, 0, data, offset, docIdData.Length);
                    offset += docIdData.Length;
                }
                var kd = md5.ComputeHash(data);
                var result = new byte[KeyDigestLength];
                Array.Copy(kd, 0, result, 0, KeyDigestLength);
                md5.Clear();

                return result;
            }
        }
        #endregion

        #region Validate
        /// <summary>
        /// Returns <c>true</c> if the keyDigest is compatible with the specified saltData and saltHash
        /// </summary>
        /// <param name="saltData"></param>
        /// <param name="saltHash"></param>
        /// <returns></returns>
        public bool Validate(byte[] saltData, byte[] saltHash)
        {
            Check16Bytes(saltData, "saltData");
            Check16Bytes(saltHash, "saltHash");

            // validation uses the RC4 for block zero
            var rc4 = CreateRC4(0);
            var saltDataPrime = new byte[saltData.Length];
            Array.Copy(saltData, saltDataPrime, saltData.Length);
            rc4.Encrypt(saltDataPrime);

            var saltHashPrime = new byte[saltHash.Length];
            Array.Copy(saltHash, saltHashPrime, saltHash.Length);
            rc4.Encrypt(saltHashPrime);

            using (MD5 md5 = new MD5CryptoServiceProvider())
            {
                var finalSaltResult = md5.ComputeHash(saltDataPrime);
                return Arrays.Equals(saltHashPrime, finalSaltResult);
            }
        }
        #endregion

        #region Check16Bytes
        private static void Check16Bytes(ICollection<byte> data, string argument)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (data.Count != 16)
                throw new ArgumentException("Expected 16 byte for " + argument);
        }
        #endregion

        #region CreateRC4
        /// <summary>
        /// The <see cref="RC4"/> instance needs to be Changed every 1024 bytes.
        /// </summary>
        /// <param name="keyBlockNo"></param>
        /// <returns></returns>
        internal RC4 CreateRC4(int keyBlockNo)
        {
            using (MD5 md5 = new MD5CryptoServiceProvider())
            {
                using (var baos = new MemoryStream(4))
                {
                    new LittleEndianOutputStream(baos).WriteInt(keyBlockNo);
                    var baosToArray = baos.ToArray();
                    var data = new byte[baosToArray.Length + _keyDigest.Length];
                    Array.Copy(_keyDigest, 0, data, 0, _keyDigest.Length);
                    Array.Copy(baosToArray, 0, data, _keyDigest.Length, baosToArray.Length);

                    var digest = md5.ComputeHash(data);
                    return new RC4(digest);
                }
            }
        }
        #endregion
    }
}
