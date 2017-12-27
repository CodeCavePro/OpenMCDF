/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using RedBlackTree;

namespace OpenMcdf
{
    internal class DirectoryEntry : IDirectoryEntry
    {
        internal const int THIS_IS_GREATER = 1;
        internal const int OTHER_IS_GREATER = -1;
        private readonly IList<IDirectoryEntry> _dirRepository;

        public int SID { get; set; } = -1;

        internal static int nostream
            = unchecked((int) 0xFFFFFFFF);

        private DirectoryEntry(string name, StgType stgType, IList<IDirectoryEntry> dirRepository)
        {
            _dirRepository = dirRepository;

            StgType = stgType;

            switch (stgType)
            {
                case StgType.StgStream:

                    _storageCLSID = new Guid("00000000000000000000000000000000");
                    CreationDate = new byte[] {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
                    ModifyDate = new byte[] {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
                    break;

                case StgType.StgStorage:
                    CreationDate = BitConverter.GetBytes((DateTime.Now.ToFileTime()));
                    break;

                case StgType.StgRoot:
                    CreationDate = new byte[] {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
                    ModifyDate = new byte[] {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
                    break;
            }

            SetEntryName(name);
        }

        public byte[] EntryName
        {
            get;
            private set;
        } = new byte[64];

        public string GetEntryName()
        {
            if (EntryName != null && EntryName.Length > 0)
            {
                return Encoding.Unicode.GetString(EntryName).Remove((_nameLength - 1) / 2);
            }
            return string.Empty;
        }

        public void SetEntryName(string entryName)
        {
            if (
                entryName.Contains(@"\") ||
                entryName.Contains(@"/") ||
                entryName.Contains(@":") ||
                entryName.Contains(@"!")
            )
                throw new CFException(
                    "Invalid character in entry: the characters '\\', '/', ':','!' cannot be used in entry name");

            if (entryName.Length > 31)
                throw new CFException("Entry name MUST be smaller than 31 characters");


            var temp = Encoding.Unicode.GetBytes(entryName);
            var newName = new byte[64];
            Buffer.BlockCopy(temp, 0, newName, 0, temp.Length);
            newName[temp.Length] = 0x00;
            newName[temp.Length + 1] = 0x00;

            EntryName = newName;
            _nameLength = (ushort) (temp.Length + 2);
        }

        private ushort _nameLength;

        public ushort NameLength
        {
            get => _nameLength;
            set => throw new NotImplementedException();
        }

        public StgType StgType { get; set; }

        public StgColor StgColor { get; set; } = StgColor.Black;

        public int LeftSibling { get; set; } = nostream;

        public int RightSibling { get; set; } = nostream;

        public int Child { get; set; } = nostream;

        private Guid _storageCLSID
            = Guid.NewGuid();

        public Guid StorageCLSID
        {
            get => _storageCLSID;
            set => _storageCLSID = value;
        }

        public int StateBits { get; set; }

        public byte[] CreationDate { get; set; } = {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};

        public byte[] ModifyDate { get; set; } = {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};

        public int StartSetc { get; set; } = Sector.ENDOFCHAIN;

        public long Size { get; set; }


        public int CompareTo(object obj)
        {
            if (!(obj is IDirectoryEntry otherDir))
                throw new CFException("Invalid casting: compared object does not implement IDirectorEntry interface");

            if (NameLength > otherDir.NameLength)
            {
                return THIS_IS_GREATER;
            }
            if (NameLength < otherDir.NameLength)
            {
                return OTHER_IS_GREATER;
            }
            var thisName = Encoding.Unicode.GetString(EntryName, 0, NameLength);
            var otherName = Encoding.Unicode.GetString(otherDir.EntryName, 0, otherDir.NameLength);

            for (var z = 0; z < thisName.Length; z++)
            {
                var thisChar = char.ToUpperInvariant(thisName[z]);
                var otherChar = char.ToUpperInvariant(otherName[z]);

                if (thisChar > otherChar)
                    return THIS_IS_GREATER;
                if (thisChar < otherChar)
                    return OTHER_IS_GREATER;
            }

            return 0;

            //   return String.Compare(Encoding.Unicode.GetString(this.EntryName).ToUpper(), Encoding.Unicode.GetString(other.EntryName).ToUpper());
        }

        public override bool Equals(object obj)
        {
            return CompareTo(obj) == 0;
        }

        /// <summary>
        /// FNV hash, short for Fowler/Noll/Vo
        /// </summary>
        /// <param name="buffer"></param>
        /// <returns>(not warranted) unique hash for byte array</returns>
        private static ulong FnvHash(byte[] buffer)
        {
            ulong h = 2166136261;
            int i;

            for (i = 0; i < buffer.Length; i++)
                h = (h * 16777619) ^ buffer[i];

            return h;
        }

        public override int GetHashCode()
        {
            // ReSharper disable once NonReadonlyMemberInGetHashCode
            return (int) FnvHash(EntryName);
        }

        public void Write(Stream stream)
        {
            var rw = new StreamRW(stream);

            rw.Write(EntryName);
            rw.Write(_nameLength);
            rw.Write((byte) StgType);
            rw.Write((byte) StgColor);
            rw.Write(LeftSibling);
            rw.Write(RightSibling);
            rw.Write(Child);
            rw.Write(_storageCLSID.ToByteArray());
            rw.Write(StateBits);
            rw.Write(CreationDate);
            rw.Write(ModifyDate);
            rw.Write(StartSetc);
            rw.Write(Size);

            rw.Close();
        }

        //public Byte[] ToByteArray()
        //{
        //    MemoryStream ms
        //        = new MemoryStream(128);

        //    BinaryWriter bw = new BinaryWriter(ms);

        //    byte[] paddedName = new byte[64];
        //    Array.Copy(entryName, paddedName, entryName.Length);

        //    bw.Write(paddedName);
        //    bw.Write(nameLength);
        //    bw.Write((byte)stgType);
        //    bw.Write((byte)stgColor);
        //    bw.Write(leftSibling);
        //    bw.Write(rightSibling);
        //    bw.Write(child);
        //    bw.Write(storageCLSID.ToByteArray());
        //    bw.Write(stateBits);
        //    bw.Write(creationDate);
        //    bw.Write(modifyDate);
        //    bw.Write(startSetc);
        //    bw.Write(size);

        //    return ms.ToArray();
        //}

        public void Read(Stream stream, CFSVersion ver = CFSVersion.Ver_3)
        {
            var rw = new StreamRW(stream);

            EntryName = rw.ReadBytes(64);
            _nameLength = rw.ReadUInt16();
            StgType = (StgType) rw.ReadByte();
            //rw.ReadByte();//Ignore color, only black tree
            StgColor = (StgColor) rw.ReadByte();
            LeftSibling = rw.ReadInt32();
            RightSibling = rw.ReadInt32();
            Child = rw.ReadInt32();

            // Thanks to bugaccount (BugTrack id 3519554)
            if (StgType == StgType.StgInvalid)
            {
                LeftSibling = nostream;
                RightSibling = nostream;
                Child = nostream;
            }

            _storageCLSID = new Guid(rw.ReadBytes(16));
            StateBits = rw.ReadInt32();
            CreationDate = rw.ReadBytes(8);
            ModifyDate = rw.ReadBytes(8);
            StartSetc = rw.ReadInt32();

            if (ver == CFSVersion.Ver_3)
            {
                // avoid dirty read for version 3 files (max size: 32bit integer)
                // where most significant bits are not initialized to zero

                Size = rw.ReadInt32();
                rw.ReadBytes(4); //discard most significant 4 (possibly) dirty bytes
            }
            else
            {
                Size = rw.ReadInt64();
            }
        }

        public string Name => GetEntryName();


        public IRbNode Left
        {
            get
            {
                if (LeftSibling == nostream)
                    return null;

                return _dirRepository[LeftSibling];
            }
            set
            {
                LeftSibling = value != null ? ((IDirectoryEntry) value).SID : nostream;

                if (LeftSibling != nostream)
                    _dirRepository[LeftSibling].Parent = this;
            }
        }

        public IRbNode Right
        {
            get
            {
                if (RightSibling == nostream)
                    return null;

                return _dirRepository[RightSibling];
            }
            set
            {
                RightSibling = value != null ? ((IDirectoryEntry) value).SID : nostream;

                if (RightSibling != nostream)
                    _dirRepository[RightSibling].Parent = this;
            }
        }

        public Color Color
        {
            get => (Color) StgColor;
            set => StgColor = (StgColor) value;
        }

        private IDirectoryEntry _parent;

        public IRbNode Parent
        {
            get => _parent;
            set => _parent = value as IDirectoryEntry;
        }

        public IRbNode Grandparent()
        {
            return _parent?.Parent;
        }

        public IRbNode Sibling()
        {
            return Equals(this, Parent.Left) ? Parent.Right : Parent.Left;
        }

        public IRbNode Uncle()
        {
            return _parent != null ? Parent.Sibling() : null;
        }

        internal static IDirectoryEntry New(string name, StgType stgType, IList<IDirectoryEntry> dirRepository)
        {
            DirectoryEntry de;
            if (dirRepository != null)
            {
                de = new DirectoryEntry(name, stgType, dirRepository);
                // No invalid directory entry found
                dirRepository.Add(de);
                de.SID = dirRepository.Count - 1;
            }
            else
                throw new ArgumentNullException("dirRepository", "Directory repository cannot be null in New() method");

            return de;
        }

        internal static IDirectoryEntry Mock(string name, StgType stgType)
        {
            var de = new DirectoryEntry(name, stgType, null);

            return de;
        }

        internal static IDirectoryEntry TryNew(string name, StgType stgType, IList<IDirectoryEntry> dirRepository)
        {
            var de = new DirectoryEntry(name, stgType, dirRepository);

            // If we are not adding an invalid dirEntry as
            // in a normal loading from file (invalid directories MAY pad a sector)
            // Find first available invalid slot (if any) to reuse it
            for (var i = 0; i < dirRepository.Count; i++)
            {
                if (dirRepository[i].StgType != StgType.StgInvalid)
                    continue;

                dirRepository[i] = de;
                de.SID = i;
                return de;
            }

            // No invalid directory entry found
            dirRepository.Add(de);
            de.SID = dirRepository.Count - 1;

            return de;
        }

        public override string ToString()
        {
            return Name + " [" + SID + "]" + (StgType == StgType.StgStream ? "Stream" : "Storage");
        }

        public void AssignValueTo(IRbNode other)
        {
            if (!(other is DirectoryEntry d))
                return;

            d.SetEntryName(GetEntryName());

            d.CreationDate = new byte[CreationDate.Length];
            CreationDate.CopyTo(d.CreationDate, 0);

            d.ModifyDate = new byte[ModifyDate.Length];
            ModifyDate.CopyTo(d.ModifyDate, 0);

            d.Size = Size;
            d.StartSetc = StartSetc;
            d.StateBits = StateBits;
            d.StgType = StgType;
            d._storageCLSID = new Guid(_storageCLSID.ToByteArray());
            d.Child = Child;
        }
    }
}