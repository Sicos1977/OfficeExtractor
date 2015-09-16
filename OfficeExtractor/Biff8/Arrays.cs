using System;
using System.Linq;

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