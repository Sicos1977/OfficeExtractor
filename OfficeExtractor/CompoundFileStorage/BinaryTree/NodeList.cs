using System.Collections.ObjectModel;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.BinaryTree
{
    /// <summary>
    ///     Represents a collection of Node&lt;T&gt; instances.
    /// </summary>
    /// <typeparam name="T">The type of data held in the Node instances referenced by this class.</typeparam>
    public class NodeList<T> : Collection<Node<T>>
    {
        #region Constructors
        public NodeList()
        {
        }

        public NodeList(int initialSize)
        {
            // Add the specified number of items
            for (var i = 0; i < initialSize; i++)
                Items.Add(default(Node<T>));
        }
        #endregion

        #region FindByValue
        /// <summary>
        ///     Searches the NodeList for a Node containing a particular value.
        /// </summary>
        /// <param name="value">The value to search for.</param>
        /// <returns>The Node in the NodeList, if it exists; null otherwise.</returns>
        public Node<T> FindByValue(T value)
        {
            // search the list for the value
            foreach (var node in Items)
                if (node.Value.Equals(value))
                    return node;

            // if we reached here, we didn't find a matching node
            return null;
        }
        #endregion
    }
}