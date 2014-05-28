
namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.BinaryTree
{
    /// <summary>
    ///     The BinaryTreeNode class represents a node in a binary tree, or a binary search tree.
    ///     It has precisely two neighbors, which can be accessed via the Left and Right properties.
    /// </summary>
    /// <typeparam name="T">The type of data stored in the binary tree node.</typeparam>
    public class BinaryTreeNode<T> : Node<T>
    {
        #region Properties
        public BinaryTreeNode<T> Left
        {
            get
            {
                if (Neighbors == null)
                    return null;
                return (BinaryTreeNode<T>) Neighbors[0];
            }
            set
            {
                if (Neighbors == null)
                    Neighbors = new NodeList<T>(2);

                Neighbors[0] = value;
            }
        }

        public BinaryTreeNode<T> Right
        {
            get
            {
                if (Neighbors == null)
                    return null;
                return (BinaryTreeNode<T>) Neighbors[1];
            }
            set
            {
                if (Neighbors == null)
                    Neighbors = new NodeList<T>(2);

                Neighbors[1] = value;
            }
        }
        #endregion

        #region Constructors
        public BinaryTreeNode()
        {
        }

        public BinaryTreeNode(T data)
            : base(data, null)
        {
        }

        public BinaryTreeNode(T data, BinaryTreeNode<T> left, BinaryTreeNode<T> right)
        {
            Value = data;
            var children = new NodeList<T>(2);
            children[0] = left;
            children[1] = right;

            Neighbors = children;
        }
        #endregion
    }
}