using System;
using System.Collections.Generic;
using NUnit.Framework;
using RedBlackTree;

namespace OpenMcdf.Test
{
    /// <summary>
    /// Summary description for RBTreeTest
    /// </summary>
    [TestFixture]
    public class RBTreeTest
    {
        internal IList<IDirectoryEntry> GetDirectoryRepository(int count)
        {
            List<IDirectoryEntry> repo = new List<IDirectoryEntry>();
            for (int i = 0; i < count; i++)
            {
                IDirectoryEntry de = DirectoryEntry.New(i.ToString(), StgType.StgInvalid, repo);
            }

            return repo;
        }

        [Test]
        public void Test_RBTREE_INSERT()
        {
            RbTree rbTree = new RbTree();
            System.Collections.Generic.IList<IDirectoryEntry> repo = GetDirectoryRepository(25);

            foreach (var item in repo)
            {
                rbTree.Insert(item);
            }

            for (int i = 0; i < repo.Count; i++)
            {
                IRbNode c;
                rbTree.TryLookup(DirectoryEntry.Mock(i.ToString(), StgType.StgInvalid), out c);
                Assert.IsTrue(c is IDirectoryEntry);
                Assert.IsTrue(((IDirectoryEntry) c).Name == i.ToString());
                //Assert.IsTrue(c.IsStream);
            }
        }


        [Test]
        public void Test_RBTREE_DELETE()
        {
            RbTree rbTree = new RbTree();
            System.Collections.Generic.IList<IDirectoryEntry> repo = GetDirectoryRepository(25);


            foreach (var item in repo)
            {
                rbTree.Insert(item);
            }

            try
            {
                IRbNode n;
                rbTree.Delete(DirectoryEntry.Mock("5", StgType.StgInvalid), out n);
                rbTree.Delete(DirectoryEntry.Mock("24", StgType.StgInvalid), out n);
                rbTree.Delete(DirectoryEntry.Mock("7", StgType.StgInvalid), out n);
            }
            catch (Exception ex)
            {
                Assert.Fail("Item removal failed: " + ex.Message);
            }


            //    CFItem c;
            //    bool s = rbTree.TryLookup(new CFMock("7", StgType.StgStream), out c);


            //    Assert.IsFalse(s);

            //    c = null;

            //    Assert.IsTrue(rbTree.TryLookup(new CFMock("6", StgType.StgStream), out c));
            //    Assert.IsTrue(c.IsStream);
            //    Assert.IsTrue(rbTree.TryLookup(new CFMock("12", StgType.StgStream), out c));
            //    Assert.IsTrue(c.Name == "12");


            //}
        }

        private static void VerifyProperties(RbTree t)
        {
            VerifyProperty1(t.Root);
            VerifyProperty2(t.Root);
            // Property 3 is implicit
            VerifyProperty4(t.Root);
            VerifyProperty5(t.Root);
        }

        private static Color NodeColor(IRbNode n)
        {
            return n == null ? Color.Black : n.Color;
        }

        private static void VerifyProperty1(IRbNode n)
        {
            Assert.IsTrue(NodeColor(n) == Color.Red || NodeColor(n) == Color.Black);

            if (n == null) return;
            VerifyProperty1(n.Left);
            VerifyProperty1(n.Right);
        }

        private static void VerifyProperty2(IRbNode root)
        {
            Assert.IsTrue(NodeColor(root) == Color.Black);
        }

        private static void VerifyProperty4(IRbNode n)
        {
            if (NodeColor(n) == Color.Red)
            {
                Assert.IsTrue((NodeColor(n.Left) == Color.Black));
                Assert.IsTrue((NodeColor(n.Right) == Color.Black));
                Assert.IsTrue((NodeColor(n.Parent) == Color.Black));
            }

            if (n == null) return;
            VerifyProperty4(n.Left);
            VerifyProperty4(n.Right);
        }

        private static void VerifyProperty5(IRbNode root)
        {
            VerifyProperty5Helper(root, 0, -1);
        }

        private static int VerifyProperty5Helper(IRbNode n, int blackCount, int pathBlackCount)
        {
            if (NodeColor(n) == Color.Black)
            {
                blackCount++;
            }
            if (n == null)
            {
                if (pathBlackCount == -1)
                {
                    pathBlackCount = blackCount;
                }
                else
                {
                    Assert.IsTrue(blackCount == pathBlackCount);
                }
                return pathBlackCount;
            }

            pathBlackCount = VerifyProperty5Helper(n.Left, blackCount, pathBlackCount);
            pathBlackCount = VerifyProperty5Helper(n.Right, blackCount, pathBlackCount);

            return pathBlackCount;
        }


        [Test]
        public void Test_RBTREE_ENUMERATE()
        {
            RbTree rbTree = new RbTree();
            System.Collections.Generic.IList<IDirectoryEntry> repo = GetDirectoryRepository(10000);

            foreach (var item in repo)
            {
                rbTree.Insert(item);
            }

            VerifyProperties(rbTree);
            //rbTree.Print();
        }
    }
}