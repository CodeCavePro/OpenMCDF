/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/

using System;
using System.Collections.Generic;
using RedBlackTree;

namespace OpenMcdf
{
    /// <summary>
    /// Action to apply to  visited items in the OLE structured storage
    /// </summary>
    /// <param name="item">Currently visited <see cref="T:OpenMcdf.CFItem">item</see></param>
    /// <example>
    /// <code>
    /// 
    /// //We assume that xls file should be a valid OLE compound file
    /// const String STORAGE_NAME = "report.xls";
    /// CompoundFile cf = new CompoundFile(STORAGE_NAME);
    ///
    /// FileStream output = new FileStream("LogEntries.txt", FileMode.Create);
    /// TextWriter tw = new StreamWriter(output);
    ///
    /// VisitedEntryAction va = delegate(CFItem item)
    /// {
    ///     tw.WriteLine(item.Name);
    /// };
    ///
    /// cf.RootStorage.VisitEntries(va, true);
    ///
    /// tw.Close();
    ///
    /// </code>
    /// </example>
    public delegate void VisitedEntryAction(CFItem item);

    /// <inheritdoc />
    /// <summary>
    /// Storage entity that acts like a logic container for streams
    /// or sub-storages in a compound file.
    /// </summary>
    public class CFStorage : CFItem
    {
        private RbTree _children;

        internal RbTree Children
        {
            get
            {
                // Lazy loading of children tree.
                if (_children == null)
                {
                    //if (this.CompoundFile.HasSourceStream)
                    //{
                    _children = LoadChildren(DirEntry.SID);
                    //}
                    //else
                    if (_children == null)
                    {
                        _children = CompoundFile.CreateNewTree();
                    }
                }

                return _children;
            }
        }


        /// <summary>
        /// Create a CFStorage using an existing directory (previously loaded).
        /// </summary>
        /// <param name="compFile">The Storage Owner - CompoundFile</param>
        /// <param name="dirEntry">An existing Directory Entry</param>
        internal CFStorage(CompoundFile compFile, IDirectoryEntry dirEntry)
            : base(compFile)
        {
            if (dirEntry == null || dirEntry.SID < 0)
                throw new CFException("Attempting to create a CFStorage using an uninitialized directory");

            DirEntry = dirEntry;
        }

        private RbTree LoadChildren(int sid)
        {
            var childrenTree = CompoundFile.GetChildrenTree(sid);
            DirEntry.Child = (childrenTree.Root as IDirectoryEntry)?.SID ?? DirectoryEntry.nostream;
            return childrenTree;
        }

        /// <summary>
        /// Create a new child stream inside the current <see cref="T:OpenMcdf.CFStorage">storage</see>
        /// </summary>
        /// <param name="streamName">The new stream name</param>
        /// <returns>The new <see cref="T:OpenMcdf.CFStream">stream</see> reference</returns>
        /// <exception cref="T:OpenMcdf.CFDuplicatedItemException">Raised when adding an item with the same name of an existing one</exception>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised when adding a stream to a closed compound file</exception>
        /// <exception cref="T:OpenMcdf.CFException">Raised when adding a stream with null or empty name</exception>
        /// <example>
        /// <code>
        /// 
        ///  String filename = "A_NEW_COMPOUND_FILE_YOU_CAN_WRITE_TO.cfs";
        ///
        ///  CompoundFile cf = new CompoundFile();
        ///
        ///  CFStorage st = cf.RootStorage.AddStorage("MyStorage");
        ///  CFStream sm = st.AddStream("MyStream");
        ///  byte[] b = Helpers.GetBuffer(220, 0x0A);
        ///  sm.SetData(b);
        ///
        ///  cf.Save(filename);
        ///  
        /// </code>
        /// </example>
        public CFStream AddStream(string streamName)
        {
            CheckDisposed();

            if (string.IsNullOrEmpty(streamName))
                throw new CFException("Stream name cannot be null or empty");


            var dirEntry =
                DirectoryEntry.TryNew(streamName, StgType.StgStream, CompoundFile.GetDirectories());

            // Add new Stream directory entry
            //cfo = new CFStream(this.CompoundFile, streamName);

            try
            {
                // Add object to Siblings tree
                Children.Insert(dirEntry);

                //... and set the root of the tree as new child of the current item directory entry
                DirEntry.Child = ((IDirectoryEntry) Children.Root).SID;
            }
            catch (RbTreeException)
            {
                CompoundFile.ResetDirectoryEntry(dirEntry.SID);

                throw new CFDuplicatedItemException("An entry with name '" + streamName +
                                                    "' is already present in storage '" + Name + "' ");
            }

            return new CFStream(CompoundFile, dirEntry);
        }


        /// <summary>
        /// Get a named <see cref="T:OpenMcdf.CFStream">stream</see> contained in the current storage if existing.
        /// </summary>
        /// <param name="streamName">Name of the stream to look for</param>
        /// <returns>A stream reference if existing</returns>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised if trying to delete item from a closed compound file</exception>
        /// <exception cref="T:OpenMcdf.CFItemNotFound">Raised if item to delete is not found</exception>
        /// <example>
        /// <code>
        /// String filename = "report.xls";
        ///
        /// CompoundFile cf = new CompoundFile(filename);
        /// CFStream foundStream = cf.RootStorage.GetStream("Workbook");
        ///
        /// byte[] temp = foundStream.GetData();
        ///
        /// Assert.IsNotNull(temp);
        ///
        /// cf.Close();
        /// </code>
        /// </example>
        public CFStream GetStream(string streamName)
        {
            CheckDisposed();

            var tmp = DirectoryEntry.Mock(streamName, StgType.StgStream);

            if (Children.TryLookup(tmp, out var outDe) && (((IDirectoryEntry) outDe).StgType == StgType.StgStream))
            {
                return new CFStream(CompoundFile, (IDirectoryEntry) outDe);
            }
            throw new CFItemNotFound("Cannot find item [" + streamName + "] within the current storage");
        }


        /// <summary>
        /// Get a named <see cref="T:OpenMcdf.CFStream">stream</see> contained in the current storage if existing.
        /// </summary>
        /// <param name="streamName">Name of the stream to look for</param>
        /// <returns>A stream reference if found, else null</returns>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised if trying to delete item from a closed compound file</exception>
        /// <example>
        /// <code>
        /// String filename = "report.xls";
        ///
        /// CompoundFile cf = new CompoundFile(filename);
        /// CFStream foundStream = cf.RootStorage.TryGetStream("Workbook");
        ///
        /// byte[] temp = foundStream.GetData();
        ///
        /// Assert.IsNotNull(temp);
        ///
        /// cf.Close();
        /// </code>
        /// </example>
        public CFStream TryGetStream(string streamName)
        {
            CheckDisposed();

            var tmp = DirectoryEntry.Mock(streamName, StgType.StgStream);

            if (Children.TryLookup(tmp, out var outDe) && (((IDirectoryEntry) outDe).StgType == StgType.StgStream))
            {
                return new CFStream(CompoundFile, (IDirectoryEntry) outDe);
            }
            return null;
        }


        /// <summary>
        /// Get a named storage contained in the current one if existing.
        /// </summary>
        /// <param name="storageName">Name of the storage to look for</param>
        /// <returns>A storage reference if existing.</returns>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised if trying to delete item from a closed compound file</exception>
        /// <exception cref="T:OpenMcdf.CFItemNotFound">Raised if item to delete is not found</exception>
        /// <example>
        /// <code>
        /// 
        /// String FILENAME = "MultipleStorage2.cfs";
        /// CompoundFile cf = new CompoundFile(FILENAME, UpdateMode.ReadOnly, false, false);
        ///
        /// CFStorage st = cf.RootStorage.GetStorage("MyStorage");
        ///
        /// Assert.IsNotNull(st);
        /// cf.Close();
        /// </code>
        /// </example>
        public CFStorage GetStorage(string storageName)
        {
            CheckDisposed();

            var template = DirectoryEntry.Mock(storageName, StgType.StgInvalid);

            if (Children.TryLookup(template, out var outDe) && ((IDirectoryEntry) outDe).StgType == StgType.StgStorage)
            {
                return new CFStorage(CompoundFile, (IDirectoryEntry) outDe);
            }
            throw new CFItemNotFound("Cannot find item [" + storageName + "] within the current storage");
        }

        /// <summary>
        /// Get a named storage contained in the current one if existing.
        /// </summary>
        /// <param name="storageName">Name of the storage to look for</param>
        /// <returns>A storage reference if found else null</returns>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised if trying to delete item from a closed compound file</exception>
        /// <example>
        /// <code>
        /// 
        /// String FILENAME = "MultipleStorage2.cfs";
        /// CompoundFile cf = new CompoundFile(FILENAME, UpdateMode.ReadOnly, false, false);
        ///
        /// CFStorage st = cf.RootStorage.TryGetStorage("MyStorage");
        ///
        /// Assert.IsNotNull(st);
        /// cf.Close();
        /// </code>
        /// </example>
        public CFStorage TryGetStorage(string storageName)
        {
            CheckDisposed();

            var template = DirectoryEntry.Mock(storageName, StgType.StgInvalid);

            if (Children.TryLookup(template, out var outDe) && ((IDirectoryEntry) outDe).StgType == StgType.StgStorage)
            {
                return new CFStorage(CompoundFile, (IDirectoryEntry) outDe);
            }
            return null;
        }


        /// <summary>
        /// Create new child storage directory inside the current storage.
        /// </summary>
        /// <param name="storageName">The new storage name</param>
        /// <returns>Reference to the new <see cref="T:OpenMcdf.CFStorage">storage</see></returns>
        /// <exception cref="T:OpenMcdf.CFDuplicatedItemException">Raised when adding an item with the same name of an existing one</exception>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised when adding a storage to a closed compound file</exception>
        /// <exception cref="T:OpenMcdf.CFException">Raised when adding a storage with null or empty name</exception>
        /// <example>
        /// <code>
        /// 
        ///  String filename = "A_NEW_COMPOUND_FILE_YOU_CAN_WRITE_TO.cfs";
        ///
        ///  CompoundFile cf = new CompoundFile();
        ///
        ///  CFStorage st = cf.RootStorage.AddStorage("MyStorage");
        ///  CFStream sm = st.AddStream("MyStream");
        ///  byte[] b = Helpers.GetBuffer(220, 0x0A);
        ///  sm.SetData(b);
        ///
        ///  cf.Save(filename);
        ///  
        /// </code>
        /// </example>
        public CFStorage AddStorage(string storageName)
        {
            CheckDisposed();

            if (string.IsNullOrEmpty(storageName))
                throw new CFException("Stream name cannot be null or empty");

            // Add new Storage directory entry
            var cfo
                = DirectoryEntry.New(storageName, StgType.StgStorage, CompoundFile.GetDirectories());

            try
            {
                // Add object to Siblings tree
                Children.Insert(cfo);
            }
            catch (RbTreeDuplicatedItemException)
            {
                CompoundFile.ResetDirectoryEntry(cfo.SID);
                throw new CFDuplicatedItemException("An entry with name '" + storageName +
                                                    "' is already present in storage '" + Name + "' ");
            }

            if (Children.Root is IDirectoryEntry childrenRoot)
                DirEntry.Child = childrenRoot.SID;

            return new CFStorage(CompoundFile, cfo);
        }

        /// <summary>
        /// Visit all entities contained in the storage applying a user provided action
        /// </summary>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised when visiting items of a closed compound file</exception>
        /// <param name="action">User <see cref="T:OpenMcdf.VisitedEntryAction">action</see> to apply to visited entities</param>
        /// <param name="recursive"> Visiting recursion level. True means substorages are visited recursively, false indicates that only the direct children of this storage are visited</param>
        /// <example>
        /// <code>
        /// const String STORAGE_NAME = "report.xls";
        /// CompoundFile cf = new CompoundFile(STORAGE_NAME);
        ///
        /// FileStream output = new FileStream("LogEntries.txt", FileMode.Create);
        /// TextWriter tw = new StreamWriter(output);
        ///
        /// VisitedEntryAction va = delegate(CFItem item)
        /// {
        ///     tw.WriteLine(item.Name);
        /// };
        ///
        /// cf.RootStorage.VisitEntries(va, true);
        ///
        /// tw.Close();
        /// </code>
        /// </example>
        public void VisitEntries(Action<CFItem> action, bool recursive)
        {
            CheckDisposed();

            if (action == null)
                return;
            var subStorages = new List<IRbNode>();

            void InternalAction(IRbNode targetNode)
            {
                var d = targetNode as IDirectoryEntry;
                if (d != null && d.StgType == StgType.StgStream)
                    action(new CFStream(CompoundFile, d));
                else
                    action(new CFStorage(CompoundFile, d));

                if (d != null && d.Child != DirectoryEntry.nostream)
                    subStorages.Add(targetNode);
            }

            Children.VisitTreeNodes(InternalAction);

            if (!recursive || subStorages.Count <= 0)
                return;

            foreach (var n in subStorages)
            {
                var d = n as IDirectoryEntry;
                (new CFStorage(CompoundFile, d)).VisitEntries(action, true);
            }
        }

        /// <summary>
        /// Remove an entry from the current storage and compound file.
        /// </summary>
        /// <param name="entryName">The name of the entry in the current storage to delete</param>
        /// <example>
        /// <code>
        /// cf = new CompoundFile("A_FILE_YOU_CAN_CHANGE.cfs", UpdateMode.Update, true, false);
        /// cf.RootStorage.Delete("AStream"); // AStream item is assumed to exist.
        /// cf.Commit(true);
        /// cf.Close();
        /// </code>
        /// </example>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised if trying to delete item from a closed compound file</exception>
        /// <exception cref="T:OpenMcdf.CFItemNotFound">Raised if item to delete is not found</exception>
        /// <exception cref="T:OpenMcdf.CFException">Raised if trying to delete root storage</exception>
        public void Delete(string entryName)
        {
            CheckDisposed();

            // Find entry to delete
            var tmp = DirectoryEntry.Mock(entryName, StgType.StgInvalid);

            Children.TryLookup(tmp, out var foundObj);

            if (foundObj == null)
                throw new CFItemNotFound("Entry named [" + entryName + "] was not found");

            if (((IDirectoryEntry) foundObj).StgType == StgType.StgRoot)
                throw new CFException("Root storage cannot be removed");

            IRbNode altDel;
            // ReSharper disable once SwitchStatementMissingSomeCases
            switch (((IDirectoryEntry) foundObj).StgType)
            {
                case StgType.StgStorage:

                    var temp = new CFStorage(CompoundFile, ((IDirectoryEntry) foundObj));

                    // This is a storage. we have to remove children items first
                    foreach (var de in temp.Children)
                    {
                        if (de is IDirectoryEntry ded) temp.Delete(ded.Name);
                    }


                    // ...then we need to re-thread the root of siblings tree...
                    DirEntry.Child = (Children.Root as IDirectoryEntry)?.SID ?? DirectoryEntry.nostream;

                    // ...and finally Remove storage item from children tree...
                    Children.Delete(foundObj, out altDel);

                    // ...and remove directory (storage) entry

                    if (altDel != null)
                    {
                        foundObj = altDel;
                    }

                    CompoundFile.InvalidateDirectoryEntry(((IDirectoryEntry) foundObj).SID);

                    break;

                case StgType.StgStream:

                    // Free directory associated data stream. 
                    CompoundFile.FreeAssociatedData((foundObj as IDirectoryEntry).SID);

                    // Remove item from children tree
                    Children.Delete(foundObj, out altDel);

                    // Re-thread the root of siblings tree...
                    DirEntry.Child = (Children.Root as IDirectoryEntry)?.SID ?? DirectoryEntry.nostream;

                    // Delete operation could possibly have cloned a directory, changing its SID.
                    // Invalidate the ACTUALLY deleted directory.
                    if (altDel != null)
                    {
                        foundObj = altDel;
                    }

                    CompoundFile.InvalidateDirectoryEntry(((IDirectoryEntry) foundObj).SID);

                    break;
            }
        }
    }
}