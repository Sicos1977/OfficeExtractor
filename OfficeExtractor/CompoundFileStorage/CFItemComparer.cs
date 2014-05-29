using System.Collections.Generic;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage
{
    internal class CFItemComparer : IComparer<CFItem>
    {
        public int Compare(CFItem x, CFItem y)
        {
            // X CompareTo Y : X > Y --> 1 ; X < Y  --> -1
            return (x.DirEntry.CompareTo(y.DirEntry));

            //Compare X < Y --> -1
        }
    }
}