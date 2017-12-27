#define ASSERT

using System;
using System.Collections.Generic;
using System.Linq;

#if ASSERT

using System.Diagnostics;

#endif

// ------------------------------------------------------------- 
// This is a porting from java code, under MIT license of       |
// the beautiful Red-Black Tree implementation you can find at  |
// http://en.literateprograms.org/Red-black_tree_(Java)#chunk   |
// Many Thanks to original Implementors.                        |
// -------------------------------------------------------------

// ReSharper disable once CheckNamespace
namespace RedBlackTree
{
    public class RbTreeException : Exception
    {
        public RbTreeException(string msg)
            : base(msg)
        {
        }
    }

    public class RbTreeDuplicatedItemException : RbTreeException
    {
        public RbTreeDuplicatedItemException(string msg)
            : base(msg)
        {
        }
    }

    public enum Color
    {
        Red = 0,
        Black = 1
    }

    /// <summary>
    /// Red Black Node interface
    /// </summary>
    public interface IRbNode : IComparable
    {
        IRbNode Left { get; set; }

        IRbNode Right { get; set; }


        Color Color { get; set; }


        IRbNode Parent { get; set; }


        IRbNode Grandparent();


        IRbNode Sibling();
        //        {
        //#if ASSERT
        //            Debug.Assert(Parent != null); // Root node has no sibling
        //#endif
        //            if (this == Parent.Left)
        //                return Parent.Right;
        //            else
        //                return Parent.Left;
        //        }

        IRbNode Uncle();
        //        {
        //#if ASSERT
        //            Debug.Assert(Parent != null); // Root node has no uncle
        //            Debug.Assert(Parent.Parent != null); // Children of root have no uncle
        //#endif
        //            return Parent.Sibling();
        //        }
        //    }

        void AssignValueTo(IRbNode other);
    }

    public class RbTree
    {
        public IRbNode Root { get; set; }

        private static Color NodeColor(IRbNode n)
        {
            return n == null ? Color.Black : n.Color;
        }

        public RbTree()
        {
        }

        public RbTree(IRbNode root)
        {
            Root = root;
        }


        private IRbNode LookupNode(IRbNode template)
        {
            var n = Root;

            while (n != null)
            {
                var compResult = template.CompareTo(n);

                if (compResult == 0)
                {
                    return n;
                }
                n = compResult < 0 ? n.Left : n.Right;
            }

            return null;
        }

        public bool TryLookup(IRbNode template, out IRbNode val)
        {
            var n = LookupNode(template);

            if (n == null)
            {
                val = null;
                return false;
            }
            val = n;
            return true;
        }

        private void ReplaceNode(IRbNode oldn, IRbNode newn)
        {
            if (oldn.Parent == null)
            {
                Root = newn;
            }
            else
            {
                if (Equals(oldn, oldn.Parent.Left))
                    oldn.Parent.Left = newn;
                else
                    oldn.Parent.Right = newn;
            }
            if (newn != null)
            {
                newn.Parent = oldn.Parent;
            }
        }

        private void RotateLeft(IRbNode n)
        {
            var r = n.Right;
            ReplaceNode(n, r);
            n.Right = r.Left;
            if (r.Left != null)
            {
                r.Left.Parent = n;
            }
            r.Left = n;
            n.Parent = r;
        }

        private void RotateRight(IRbNode n)
        {
            var l = n.Left;
            ReplaceNode(n, l);
            n.Left = l.Right;

            if (l.Right != null)
            {
                l.Right.Parent = n;
            }

            l.Right = n;
            n.Parent = l;
        }


        public void Insert(IRbNode newNode)
        {
            newNode.Color = Color.Red;
            var insertedNode = newNode;

            if (Root == null)
            {
                Root = insertedNode;
            }
            else
            {
                var n = Root;
                while (true)
                {
                    var compResult = newNode.CompareTo(n);
                    if (compResult == 0)
                    {
                        throw new RbTreeDuplicatedItemException(
                            "RBNode " + newNode + " already present in tree");
                    }
                    if (compResult < 0)
                    {
                        if (n.Left == null)
                        {
                            n.Left = insertedNode;

                            break;
                        }
                        n = n.Left;
                    }
                    else
                    {
                        //assert compResult > 0;
                        if (n.Right == null)
                        {
                            n.Right = insertedNode;

                            break;
                        }
                        n = n.Right;
                    }
                }
                insertedNode.Parent = n;
            }

            InsertCase1(insertedNode);

            NodeInserted?.Invoke(insertedNode);
        }

        //------------------------------------
        private void InsertCase1(IRbNode n)
        {
            if (n.Parent == null)
                n.Color = Color.Black;
            else
                InsertCase2(n);
        }

        //-----------------------------------
        private void InsertCase2(IRbNode n)
        {
            if (NodeColor(n.Parent) == Color.Black)
                return; // Tree is still valid
            InsertCase3(n);
        }

        //----------------------------
        private void InsertCase3(IRbNode n)
        {
            if (NodeColor(n.Uncle()) == Color.Red)
            {
                n.Parent.Color = Color.Black;
                n.Uncle().Color = Color.Black;
                n.Grandparent().Color = Color.Red;
                InsertCase1(n.Grandparent());
            }
            else
            {
                InsertCase4(n);
            }
        }

        //----------------------------
        private void InsertCase4(IRbNode n)
        {
            if (Equals(n, n.Parent.Right) && Equals(n.Parent, n.Grandparent().Left))
            {
                RotateLeft(n.Parent);
                n = n.Left;
            }
            else if (Equals(n, n.Parent.Left) && Equals(n.Parent, n.Grandparent().Right))
            {
                RotateRight(n.Parent);
                n = n.Right;
            }

            InsertCase5(n);
        }

        //----------------------------
        private void InsertCase5(IRbNode n)
        {
            n.Parent.Color = Color.Black;
            n.Grandparent().Color = Color.Red;
            if (Equals(n, n.Parent.Left) && Equals(n.Parent, n.Grandparent().Left))
            {
                RotateRight(n.Grandparent());
            }
            else
            {
                //assert n == n.parent.right && n.parent == n.grandparent().right;
                RotateLeft(n.Grandparent());
            }
        }

        private static IRbNode MaximumNode(IRbNode n)
        {
            //assert n != null;
            while (n.Right != null)
            {
                n = n.Right;
            }

            return n;
        }


        public void Delete(IRbNode template, out IRbNode deletedAlt)
        {
            deletedAlt = null;
            var n = LookupNode(template);
            if (n == null)
                return; // Key not found, do nothing
            if (n.Left != null && n.Right != null)
            {
                // Copy key/value from predecessor and then delete it instead
                var pred = MaximumNode(n.Left);
                pred.AssignValueTo(n);
                n = pred;
                deletedAlt = pred;
            }

            //assert n.left == null || n.right == null;
            var child = n.Right ?? n.Left;
            if (NodeColor(n) == Color.Black)
            {
                n.Color = NodeColor(child);
                DeleteCase1(n);
            }

            ReplaceNode(n, child);

            if (NodeColor(Root) == Color.Red)
            {
                Root.Color = Color.Black;
            }
        }

        private void DeleteCase1(IRbNode n)
        {
            if (n.Parent == null)
                return;
            DeleteCase2(n);
        }


        private void DeleteCase2(IRbNode n)
        {
            if (NodeColor(n.Sibling()) == Color.Red)
            {
                n.Parent.Color = Color.Red;
                n.Sibling().Color = Color.Black;
                if (Equals(n, n.Parent.Left))
                    RotateLeft(n.Parent);
                else
                    RotateRight(n.Parent);
            }

            DeleteCase3(n);
        }

        private void DeleteCase3(IRbNode n)
        {
            if (NodeColor(n.Parent) == Color.Black &&
                NodeColor(n.Sibling()) == Color.Black &&
                NodeColor(n.Sibling().Left) == Color.Black &&
                NodeColor(n.Sibling().Right) == Color.Black)
            {
                n.Sibling().Color = Color.Red;
                DeleteCase1(n.Parent);
            }
            else
                DeleteCase4(n);
        }

        private void DeleteCase4(IRbNode n)
        {
            if (NodeColor(n.Parent) == Color.Red &&
                NodeColor(n.Sibling()) == Color.Black &&
                NodeColor(n.Sibling().Left) == Color.Black &&
                NodeColor(n.Sibling().Right) == Color.Black)
            {
                n.Sibling().Color = Color.Red;
                n.Parent.Color = Color.Black;
            }
            else
                DeleteCase5(n);
        }

        private void DeleteCase5(IRbNode n)
        {
            if (Equals(n, n.Parent.Left) &&
                NodeColor(n.Sibling()) == Color.Black &&
                NodeColor(n.Sibling().Left) == Color.Red &&
                NodeColor(n.Sibling().Right) == Color.Black)
            {
                n.Sibling().Color = Color.Red;
                n.Sibling().Left.Color = Color.Black;
                RotateRight(n.Sibling());
            }
            else if (Equals(n, n.Parent.Right) &&
                     NodeColor(n.Sibling()) == Color.Black &&
                     NodeColor(n.Sibling().Right) == Color.Red &&
                     NodeColor(n.Sibling().Left) == Color.Black)
            {
                n.Sibling().Color = Color.Red;
                n.Sibling().Right.Color = Color.Black;
                RotateLeft(n.Sibling());
            }

            DeleteCase6(n);
        }

        private void DeleteCase6(IRbNode n)
        {
            n.Sibling().Color = NodeColor(n.Parent);
            n.Parent.Color = Color.Black;
            if (Equals(n, n.Parent.Left))
            {
                //assert nodeColor(n.sibling().right) == Color.RED;
                n.Sibling().Right.Color = Color.Black;
                RotateLeft(n.Parent);
            }
            else
            {
                //assert nodeColor(n.sibling().left) == Color.RED;
                n.Sibling().Left.Color = Color.Black;
                RotateRight(n.Parent);
            }
        }

        public void VisitTree(Action<IRbNode> action)
        {
            //IN Order visit
            var walker = Root;

            if (walker != null)
                DoVisitTree(action, walker);
        }

        private static void DoVisitTree(Action<IRbNode> action, IRbNode walker)
        {
            if (walker.Left != null)
            {
                DoVisitTree(action, walker.Left);
            }

            action?.Invoke(walker);

            if (walker.Right != null)
            {
                DoVisitTree(action, walker.Right);
            }
        }

        internal void VisitTreeNodes(Action<IRbNode> action)
        {
            //IN Order visit
            var walker = Root;

            if (walker != null)
                DoVisitTreeNodes(action, walker);
        }

        private static void DoVisitTreeNodes(Action<IRbNode> action, IRbNode walker)
        {
            if (walker.Left != null)
            {
                DoVisitTreeNodes(action, walker.Left);
            }

            action?.Invoke(walker);

            if (walker.Right != null)
            {
                DoVisitTreeNodes(action, walker.Right);
            }
        }

        public class RbTreeEnumerator : IEnumerator<IRbNode>
        {
            int _position = -1;

            private readonly Queue<IRbNode> _heap = new Queue<IRbNode>();

            internal RbTreeEnumerator(RbTree tree)
            {
                tree.VisitTreeNodes(item => _heap.Enqueue(item));
            }

            public IRbNode Current => _heap.ElementAt(_position);

            public void Dispose()
            {
            }

            object System.Collections.IEnumerator.Current => _heap.ElementAt(_position);

            public bool MoveNext()
            {
                _position++;
                return (_position < _heap.Count);
            }

            public void Reset()
            {
                _position = -1;
            }
        }

        public RbTreeEnumerator GetEnumerator()
        {
            return new RbTreeEnumerator(this);
        }

        private static int _indentStep = 15;

        public void Print()
        {
            PrintHelper(Root, 0);
        }

        private static void PrintHelper(IRbNode n, int indent)
        {
            if (n == null)
            {
                Trace.WriteLine("<empty tree>");
                return;
            }

            if (n.Left != null)
            {
                PrintHelper(n.Left, indent + _indentStep);
            }

            for (var i = 0; i < indent; i++)
                Trace.Write(" ");
            if (n.Color == Color.Black)
                Trace.WriteLine(" " + n + " ");
            else
                Trace.WriteLine("<" + n + ">");

            if (n.Right != null)
            {
                PrintHelper(n.Right, indent + _indentStep);
            }
        }

        internal void FireNodeOperation(IRbNode node, NodeOperation operation)
        {
            if (NodeOperation != null)
                NodeOperation(node, operation);
        }

        //internal void FireValueAssigned(RBNode<V> node, V value)
        //{
        //    if (ValueAssignedAction != null)
        //        ValueAssignedAction(node, value);
        //}

        internal event Action<IRbNode> NodeInserted;

        //internal event Action<RBNode<V>> NodeDeleted;
        internal event Action<IRbNode, NodeOperation> NodeOperation;
    }

    internal enum NodeOperation
    {
        LeftAssigned,
        RightAssigned,
        ColorAssigned,
        ParentAssigned,
        ValueAssigned
    }
}