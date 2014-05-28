using System;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.RBTree
{
    /// <summary>
    ///     The RedBlackNode class encapsulates a node in the tree
    /// </summary>
    public class RedBlackNode
    {
        #region Fields
        // Tree node colors
        public static int Red = 0;
        public static int Black = 1;
        #endregion

        #region Properties
        /// <summary>
        ///     Key
        /// </summary>
        public IComparable Key { get; set; }

        /// <summary>
        ///     Data
        /// </summary>
        public object Data { get; set; }

        /// <summary>
        ///     Color
        /// </summary>
        public int Color { get; set; }

        /// <summary>
        ///     Left
        /// </summary>
        public RedBlackNode Left { get; set; }

        /// <summary>
        ///     Right
        /// </summary>
        public RedBlackNode Right { get; set; }

        public RedBlackNode Parent { get; set; }
        #endregion

        #region Constructor
        public RedBlackNode()
        {
            Color = Red;
        }
        #endregion
    }
}