using System;
using System.Collections;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.RBTree.Exceptions;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.RBTree
{
    /// <summary>
    ///     The RedBlackEnumerator class returns the keys or data objects of the treap in
    ///     sorted order.
    /// </summary>
    public class RedBlackEnumerator
    {
        #region Fields
        /// <summary>
        /// The treap uses the stack to order the nodes
        // return in ascending order (true) or descending (false)
        /// </summary>
        private readonly bool _ascending;
        private readonly bool _keys;
        private readonly Stack _stack;
        public string Color; // testing only, don't use in live system
        public IComparable ParentKey; // testing only, don't use in live system
        #endregion

        #region Properties
        /// <summary>
        ///     Key
        /// </summary>
        public IComparable Key { get; set; }

        /// <summary>
        ///     Data
        /// </summary>
        public object Value { get; set; }
        #endregion

        #region Constructor
        public RedBlackEnumerator()
        {
        }
        #endregion
        
        #region RedBlackEnumerator
        /// <summary>
        ///     Determine order, walk the tree and push the nodes onto the stack
        /// </summary>
        public RedBlackEnumerator(RedBlackNode tnode, bool keys, bool ascending)
        {
            _stack = new Stack();
            _keys = keys;
            _ascending = ascending;

            // use depth-first traversal to push nodes into stack
            // the lowest node will be at the top of the stack
            if (ascending)
            {
                // find the lowest node
                while (tnode != RedBlack.SentinelNode)
                {
                    _stack.Push(tnode);
                    tnode = tnode.Left;
                }
            }
            else
            {
                // the highest node will be at top of stack
                while (tnode != RedBlack.SentinelNode)
                {
                    _stack.Push(tnode);
                    tnode = tnode.Right;
                }
            }
        }
        #endregion
        
        #region HasMoreElements
        /// <summary>
        ///     HasMoreElements
        /// </summary>
        public bool HasMoreElements()
        {
            return (_stack.Count > 0);
        }
        #endregion

        #region NextElement
        /// <summary>
        ///     NextElement
        /// </summary>
        public object NextElement()
        {
            if (_stack.Count == 0)
                throw (new RedBlackException("Element not found"));

            // the top of stack will always have the next item
            // get top of stack but don't remove it as the next nodes in sequence
            // may be pushed onto the top
            // the stack will be popped after all the nodes have been returned
            var node = (RedBlackNode) _stack.Peek(); //next node in sequence

            if (_ascending)
            {
                if (node.Right == RedBlack.SentinelNode)
                {
                    // yes, top node is lowest node in subtree - pop node off stack 
                    var tn = (RedBlackNode) _stack.Pop();
                    // peek at right node's parent 
                    // get rid of it if it has already been used
                    while (HasMoreElements() && ((RedBlackNode) _stack.Peek()).Right == tn)
                        tn = (RedBlackNode) _stack.Pop();
                }
                else
                {
                    // find the next items in the sequence
                    // traverse to left; find lowest and push onto stack
                    var tn = node.Right;
                    while (tn != RedBlack.SentinelNode)
                    {
                        _stack.Push(tn);
                        tn = tn.Left;
                    }
                }
            }
            else // descending, same comments as above apply
            {
                if (node.Left == RedBlack.SentinelNode)
                {
                    // walk the tree
                    var tn = (RedBlackNode) _stack.Pop();
                    while (HasMoreElements() && ((RedBlackNode) _stack.Peek()).Left == tn)
                        tn = (RedBlackNode) _stack.Pop();
                }
                else
                {
                    // determine next node in sequence
                    // traverse to left subtree and find greatest node - push onto stack
                    var tn = node.Left;
                    while (tn != RedBlack.SentinelNode)
                    {
                        _stack.Push(tn);
                        tn = tn.Right;
                    }
                }
            }

            // the following is for .NET compatibility (see MoveNext())
            Key = node.Key;
            Value = node.Data;
            // ******** testing only ********
            try
            {
                ParentKey = node.Parent.Key; // testing only
            }
            catch (Exception)
            {
                ParentKey = 0;
            }

            Color = node.Color == 0 ? "Red" : "Black";
            // ******** testing only ********

            return _keys ? node.Key : node.Data;
        }
        #endregion

        #region MoveNext
        /// <summary>
        ///     MoveNext
        ///     For .NET compatibility
        /// </summary>
        public bool MoveNext()
        {
            if (!HasMoreElements()) return false;
            NextElement();
            return true;
        }
        #endregion

    }
}