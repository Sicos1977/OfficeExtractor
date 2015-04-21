using System;
using System.Linq;

namespace OfficeExtractor.Biff8
{
    /// <summary>
    ///     Excel array helper methods
    /// </summary>
    internal class Arrays
    {
        #region Equals
        /// <summary>
        ///     Returns true is both objects are the same
        /// </summary>
        /// <param name="a1">The a1.</param>
        /// <param name="b1">The b1.</param>
        /// <returns></returns>
        public new static bool Equals(object a1, object b1)
        {
            if (a1 == null || b1 == null)
                return false;
            var a = a1 as Array;
            var b = b1 as Array;
            // ReSharper disable PossibleNullReferenceException
            if (a.Length != b.Length)
            // ReSharper restore PossibleNullReferenceException
                return false;

            return !a.Cast<object>().Where((t, i) => !a.GetValue(i).Equals(b.GetValue(i))).Any();
        }
        #endregion
    }
}