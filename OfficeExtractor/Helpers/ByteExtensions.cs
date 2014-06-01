using System;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.Helpers
{
    /// <summary>
    /// Contains extension methods for byte types
    /// </summary>
    internal static class ByteExtensions
    {
        #region GetBit
        /// <summary>
        /// Gets the state of a bit in a byte (Assume 0 is the MSB andd 7 is the LSB)
        /// </summary>
        /// <param name="byt">The byte to read from</param>
        /// <param name="position">The position of the bit (zero based)</param>
        /// <returns></returns>

        public static bool GetBit(this byte byt, int position)
        {
            if (position < 0 || position > 7)
                throw new ArgumentOutOfRangeException();

            var shift = 7 - position;

            // Get a single bit in the proper position.
            var bitMask = (byte) (1 << shift);

            // Mask out the appropriate bit.
            var masked = (byte) (byt & bitMask);

            // If masked != 0, then the masked out bit is 1.
            // Otherwise, masked will be 0.
            return masked != 0;
        }
        #endregion
    }
}
