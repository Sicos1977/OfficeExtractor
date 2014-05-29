using System;
using System.Collections;
using System.Collections.Generic;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.BinaryTree.Exceptions;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.BinaryTree
{
    public delegate void NodeAction<T>(BinaryTreeNode<T> node);

    /// <summary>
    ///     Represents a binary search tree. A binary search tree is a binary tree whose nodes are arranged
    ///     such that for any given node k, all nodes in k's left subtree have a value less than k, and all
    ///     nodes in k's right subtree have a value greater than k.
    /// </summary>
    /// <typeparam name="T">The type of data stored in the binary tree nodes.</typeparam>
    public class BinarySearchTree<T> : ICollection<T>
    {
        #region Fields
        /// <summary>
        ///     Used to compare node values when percolating down the tree
        /// </summary>
        private readonly IComparer<T> _comparer = Comparer<T>.Default;
        #endregion

        #region Properties
        public BinaryTreeNode<T> Root { get; private set; }

        /// <summary>
        ///     Provides enumeration through the BST using preorder traversal.
        /// </summary>
        public IEnumerable<T> Preorder
        {
            get
            {
                // A single stack is sufficient here - it simply maintains the correct
                // order with which to process the children.
                var toVisit = new Stack<BinaryTreeNode<T>>(Count);
                BinaryTreeNode<T> current = Root;
                if (current != null) toVisit.Push(current);

                while (toVisit.Count != 0)
                {
                    // take the top item from the stack
                    current = toVisit.Pop();

                    // add the right and left children, if not null
                    if (current.Right != null) toVisit.Push(current.Right);
                    if (current.Left != null) toVisit.Push(current.Left);

                    // return the current node
                    yield return current.Value;
                }
            }
        }

        /// <summary>
        ///     Provides enumeration through the BST using inorder traversal.
        /// </summary>
        public IEnumerable<T> Inorder
        {
            get
            {
                // A single stack is sufficient - this code was made available by Grant Richins:
                // http://blogs.msdn.com/grantri/archive/2004/04/08/110165.aspx
                var toVisit = new Stack<BinaryTreeNode<T>>(Count);
                for (BinaryTreeNode<T> current = Root; current != null || toVisit.Count != 0; current = current.Right)
                {
                    // Get the left-most item in the subtree, remembering the path taken
                    while (current != null)
                    {
                        toVisit.Push(current);
                        current = current.Left;
                    }

                    current = toVisit.Pop();
                    yield return current.Value;
                }
            }
        }

        /// <summary>
        ///     Provides enumeration through the BST using postorder traversal.
        /// </summary>
        public IEnumerable<T> Postorder
        {
            get
            {
                // Maintain two stacks, one of a list of nodes to visit,
                // and one of booleans, indicating if the note has been processed
                // or not.
                var toVisit = new Stack<BinaryTreeNode<T>>(Count);
                var hasBeenProcessed = new Stack<bool>(Count);
                BinaryTreeNode<T> current = Root;

                if (current != null)
                {
                    toVisit.Push(current);
                    hasBeenProcessed.Push(false);
                    current = current.Left;
                }

                while (toVisit.Count != 0)
                {
                    if (current != null)
                    {
                        // Add this node to the stack with a false processed value
                        toVisit.Push(current);
                        hasBeenProcessed.Push(false);
                        current = current.Left;
                    }
                    else
                    {
                        // See if the node on the stack has been processed
                        bool processed = hasBeenProcessed.Pop();
                        BinaryTreeNode<T> node = toVisit.Pop();
                        if (!processed)
                        {
                            // If it's not been processed, "recurse" down the right subtree
                            toVisit.Push(node);
                            hasBeenProcessed.Push(true); // it's now been processed
                            current = node.Right;
                        }
                        else
                            yield return node.Value;
                    }
                }
            }
        }

        /// <summary>
        ///     Returns the number of elements in the BST.
        /// </summary>
        public int Count { get; private set; }

        public bool IsReadOnly
        {
            get { return false; }
        }
        #endregion

        #region Constructors
        public BinarySearchTree()
        {
        }

        public BinarySearchTree(IComparer<T> comparer)
        {
            _comparer = comparer;
        }
        #endregion

        #region Clear
        /// <summary>
        ///     Removes the contents of the BST
        /// </summary>
        public void Clear()
        {
            Root = null;
            Count = 0;
        }
        #endregion

        #region CopyTo
        /// <summary>
        ///     Copies the contents of the BST to an appropriately-sized array of type T, using the Inorder
        ///     traversal method.
        /// </summary>
        public void CopyTo(T[] array, int index)
        {
            CopyTo(array, index, TraversalMethod.Inorder);
        }

        /// <summary>
        ///     Copies the contents of the BST to an appropriately-sized array of type T, using a specified
        ///     traversal method.
        /// </summary>
        public void CopyTo(T[] array, int index, TraversalMethod traversalMethod)
        {
            IEnumerable<T> enumProp;

            // Determine which Enumerator-returning property to use, based on the TraversalMethod input parameter
            switch (traversalMethod)
            {
                case TraversalMethod.Preorder:
                    enumProp = Preorder;
                    break;

                case TraversalMethod.Inorder:
                    enumProp = Inorder;
                    break;

                default:
                    enumProp = Postorder;
                    break;
            }

            // dump the contents of the tree into the passed-in array
            int i = 0;
            foreach (T value in enumProp)
            {
                array[i + index] = value;
                i++;
            }
        }
        #endregion

        #region Add
        /// <summary>
        ///     Adds a new value to the BST.
        /// </summary>
        /// <param name="data">The data to insert into the BST.</param>
        /// <remarks>
        ///     Adding a value already in the BST has no effect; that is, the SkipList is not
        ///     altered, the Add() method simply exits.
        /// </remarks>
        public virtual void Add(T data)
        {
            // create a new Node instance
            var node = new BinaryTreeNode<T>(data);
            int result;

            // now, insert n into the tree
            // trace down the tree until we hit a NULL
            BinaryTreeNode<T> current = Root, parent = null;
            while (current != null)
            {
                result = _comparer.Compare(current.Value, data);
                if (result == 0)
                    // they are equal - attempting to enter a duplicate - do nothing
                    throw new BSTDuplicatedException();

                if (result < 0)
                {
                    // current.Value < data, must add node to current's right subtree
                    parent = current;
                    current = current.Right;
                }
                else if (result > 0)
                {
                    // current.Value > data, must add node to current's left subtree
                    parent = current;
                    current = current.Left;
                }
            }

            // We're ready to add the node!
            Count++;

            if (parent == null)
                // the tree was empty, make n the root
                Root = node;
            else
            {
                result = _comparer.Compare(parent.Value, data);
                if (result > 0)
                    // parent.Value > data , therefore node must be added to the left subtree
                    parent.Left = node;
                else
                    // parent.Value < data , therefore node must be added to the right subtree
                    parent.Right = node;
            }
        }
        #endregion

        #region Contains
        /// <summary>
        ///     Returns a Boolean, indicating if a specified value is contained within the BST.
        /// </summary>
        /// <param name="data">The data to search for.</param>
        /// <returns>True if data is found in the BST; false otherwise.</returns>
        public bool Contains(T data)
        {
            // search the tree for a node that contains data
            BinaryTreeNode<T> current = Root;

            while (current != null)
            {
                int result = _comparer.Compare(current.Value, data);
                if (result == 0)
                    // we found data
                    return true;

                if (result > 0)
                    // current.Value > data, search current's left subtree
                    current = current.Left;
                else if (result < 0)
                    // current.Value < data, search current's right subtree
                    current = current.Right;
            }

            return false;
        }
        #endregion

        #region TryFind
        /// <summary>
        ///     Returns a boolean, indicating if a specified value is contained within the BST.
        /// </summary>
        /// <param name="data">The data to search for.</param>
        /// <param name="foundObject"></param>
        /// <returns>True if data is found in the BST; false otherwise.</returns>
        public bool TryFind(T data, out T foundObject)
        {
            foundObject = default(T);

            // Search the tree for a node that contains data
            var current = Root;
            while (current != null)
            {
                var result = _comparer.Compare(current.Value, data);
                if (result == 0)
                {
                    // We found data
                    foundObject = current.Value;
                    return true;
                }

                if (result > 0)
                    // current.Value > data, search current's right subtree
                    current = current.Left;
                else if (result < 0)
                    // current.Value < data, search current's left subtree
                    current = current.Right;
            }

            return false;
        }
        #endregion

        #region Remove
        /// <summary>
        ///     Attempts to remove the specified data element from the BST.
        /// </summary>
        /// <param name="data">The data to remove from the BST.</param>
        /// <returns>
        ///     True if the element is found in the tree, and removed; false if the element is not
        ///     found in the tree.
        /// </returns>
        public bool Remove(T data)
        {
            // first make sure there exist some items in this tree
            if (Root == null)
                return false; // no items to remove

            // Now, try to find data in the tree
            BinaryTreeNode<T> current = Root, parent = null;
            int result = _comparer.Compare(current.Value, data);
            while (result != 0)
            {
                if (result > 0)
                {
                    // current.Value > data, if data exists it's in the left subtree
                    parent = current;
                    current = current.Left;
                }
                else if (result < 0)
                {
                    // current.Value < data, if data exists it's in the right subtree
                    parent = current;
                    current = current.Right;
                }

                // If current == null, then we didn't find the item to remove
                if (current == null)
                    return false;

                result = _comparer.Compare(current.Value, data);
            }

            // At this point, we've found the node to remove
            Count--;

            // We now need to "rethread" the tree
            // CASE 1: If current has no right child, then current's left child becomes
            //         the node pointed to by the parent
            if (current.Right == null)
            {
                if (parent == null)
                    Root = current.Left;
                else
                {
                    result = _comparer.Compare(parent.Value, current.Value);
                    if (result > 0)
                        // parent.Value > current.Value, so make current's left child a left child of parent
                        parent.Left = current.Left;
                    else if (result < 0)
                        // parent.Value < current.Value, so make current's left child a right child of parent
                        parent.Right = current.Left;
                }
            }
                // CASE 2: If current's right child has no left child, then current's right child
                //         replaces current in the tree
            else if (current.Right.Left == null)
            {
                current.Right.Left = current.Left;

                if (parent == null)
                    Root = current.Right;
                else
                {
                    result = _comparer.Compare(parent.Value, current.Value);
                    if (result > 0)
                        // parent.Value > current.Value, so make current's right child a left child of parent
                        parent.Left = current.Right;
                    else if (result < 0)
                        // parent.Value < current.Value, so make current's right child a right child of parent
                        parent.Right = current.Right;
                }
            }
                // CASE 3: If current's right child has a left child, replace current with current's
                //          right child's left-most descendent
            else
            {
                // We first need to find the right node's left-most child
                BinaryTreeNode<T> leftmost = current.Right.Left, lmParent = current.Right;
                while (leftmost.Left != null)
                {
                    lmParent = leftmost;
                    leftmost = leftmost.Left;
                }

                // the parent's left subtree becomes the leftmost's right subtree
                lmParent.Left = leftmost.Right;

                // assign leftmost's left and right to current's left and right children
                leftmost.Left = current.Left;
                leftmost.Right = current.Right;

                if (parent == null)
                    Root = leftmost;
                else
                {
                    result = _comparer.Compare(parent.Value, current.Value);
                    if (result > 0)
                        // parent.Value > current.Value, so make leftmost a left child of parent
                        parent.Left = leftmost;
                    else if (result < 0)
                        // parent.Value < current.Value, so make leftmost a right child of parent
                        parent.Right = leftmost;
                }
            }

            // Clear out the values from current
            current.Left = current.Right = null;

            return true;
        }
        #endregion

        #region GetEnumerator
        /// <summary>
        ///     Enumerates the BST's contents using inorder traversal.
        /// </summary>
        /// <returns>An enumerator that provides inorder access to the BST's elements.</returns>
        public virtual IEnumerator<T> GetEnumerator()
        {
            return GetEnumerator(TraversalMethod.Inorder);
        }

        /// <summary>
        ///     Enumerates the BST's contents using a specified traversal method.
        /// </summary>
        /// <param name="traversalMethod">The type of traversal to perform.</param>
        /// <returns>An enumerator that provides access to the BST's elements using a specified traversal technique.</returns>
        public virtual IEnumerator<T> GetEnumerator(TraversalMethod traversalMethod)
        {
            // The traversal approaches are defined as public properties in the BST class...
            // This method simply returns the appropriate property.
            switch (traversalMethod)
            {
                case TraversalMethod.Preorder:
                    return Preorder.GetEnumerator();

                case TraversalMethod.Inorder:
                    return Inorder.GetEnumerator();

                default:
                    return Postorder.GetEnumerator();
            }
        }
        #endregion

        #region VisitTreeInOrder
        public void VisitTreeInOrder(NodeAction<T> nodeCallback)
        {
            if (Root != null)
                DoVisitInOrder(Root, nodeCallback);
        }

        public void VisitTreeInOrder(NodeAction<T> nodeCallback, bool recursive)
        {
            if (Root != null)
                DoVisitInOrder(Root, nodeCallback);
        }

        private void DoVisitInOrder(BinaryTreeNode<T> node, NodeAction<T> nodeCallback)
        {
            if (node.Left != null) DoVisitInOrder(node.Left, nodeCallback);

            if (nodeCallback != null)
            {
                nodeCallback(node);
            }

            if (node.Right != null) DoVisitInOrder(node.Right, nodeCallback);
        }
        #endregion

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }
}