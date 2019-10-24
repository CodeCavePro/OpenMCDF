/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * The Original Code is OpenMCDF - Compound Document Format library.
 *
 * The Initial Developer of the Original Code is Federico Blaseotto.*/

#define FLAT_WRITE // No optimization on the number of write operations

#pragma warning disable DF0010

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using RedBlackTree;

namespace OpenMcdf
{
    /// <inheritdoc />
    /// <summary>
    /// Standard Microsoft© Compound File implementation.
    /// It is also known as OLE/COM structured storage
    /// and contains a hierarchy of storage and stream objects providing
    /// efficient storage of multiple kinds of documents in a single file.
    /// Version 3 and 4 of specifications are supported.
    /// </summary>
    public class CompoundFile : IDisposable
    {
        #region Constants

        /// <summary>
        /// Number of DIFAT entries in the header
        /// </summary>
        private const int HEADER_DIFAT_ENTRIES_COUNT = 109;

        /// <summary>
        /// Sector ID Size (int)
        /// </summary>
        private const int SIZE_OF_SID = 4;

        /// <summary>
        /// Initial capacity of the flushing queue used
        /// to optimize commit writing operations
        /// </summary>
        private const int FLUSHING_QUEUE_SIZE = 6000;

        /// <summary>
        /// Maximum size of the flushing buffer used
        /// to optimize commit writing operations
        /// </summary>
        private const int FLUSHING_BUFFER_MAX_SIZE = 1024 * 1024 * 16;

        #endregion Constants

        /// <summary>
        /// Flag for sector recycling.
        /// </summary>
        private readonly bool _sectorRecycle;

        /// <summary>
        /// Flag for unallocated sector zeroing out.
        /// </summary>
        private readonly bool _eraseFreeSectors;

        /// <summary>
        /// Number of FAT entries in a DIFAT Sector
        /// </summary>
        private readonly int _difatSectorFATEntriesCount;

        /// <summary>
        /// Sectors ID entries in a FAT Sector
        /// </summary>
        private readonly int _fatSectorEntriesCount;

        /// <summary>
        /// The collection of file sectors.
        /// </summary>
        private SectorCollection _sectors = new SectorCollection();

        /// <summary>
        /// CompoundFile header
        /// </summary>
        private Header _header;

        private List<IDirectoryEntry> _directoryEntries = new List<IDirectoryEntry>();

        private readonly List<int> _levelSiDs = new List<int>();

        private readonly object _lockObject = new object();

        internal bool TransactionLockAdded { get; private set; }

        internal int LockSectorId { get; set; } = -1;

        internal bool TransactionLockAllocated { get; private set; }

        /// <summary>
        /// Compound underlying stream. Null when new CF has been created.
        /// </summary>
        internal Stream SourceStream { get; private set; }

        /// <summary>
        /// Gets the configuration parameters of the CompoundFile object.
        /// </summary>
        public CFSConfiguration Configuration { get; }

        /// <summary>
        /// Gets the update mode of the CompoundFile object.
        /// </summary>
        public CFSUpdateMode UpdateMode { get; }

        /// <summary>
        /// The file name (currently being read or saved)
        /// </summary>
        public string FileName { get; private set; }

        /// <summary>
        /// Gets a value indicating whether [validation exception enabled].
        /// </summary>
        /// <value>
        ///   <c>true</c> if [validation exception enabled]; otherwise, <c>false</c>.
        /// </value>
        public bool ValidationExceptionEnabled { get; } = true;

        internal bool IsClosed { get; private set; }

        internal IDirectoryEntry RootEntry => _directoryEntries[0];

        public CFSVersion Version => (CFSVersion)_header.MajorVersion;

        /// <summary>
        /// The entry point object that represents the
        /// root of the structures tree to get or set storage or
        /// stream data.
        /// </summary>
        /// <example>
        /// <code>
        ///
        ///    //Create a compound file
        ///    string FILENAME = "MyFileName.cfs";
        ///    CompoundFile ncf = new CompoundFile();
        ///
        ///    CFStorage l1 = ncf.RootStorage.AddStorage("Storage Level 1");
        ///
        ///    l1.AddStream("l1ns1");
        ///    l1.AddStream("l1ns2");
        ///    l1.AddStream("l1ns3");
        ///    CFStorage l2 = l1.AddStorage("Storage Level 2");
        ///    l2.AddStream("l2ns1");
        ///    l2.AddStream("l2ns2");
        ///
        ///    ncf.Save(FILENAME);
        ///    ncf.Close();
        /// </code>
        /// </example>
        public CFStorage RootStorage { get; private set; }

        /// <summary>
        /// Create a blank, version 3 compound file.
        /// Sector recycle is turned off to achieve the best reading/writing
        /// performance in most common scenarios.
        /// </summary>
        /// <example>
        /// <code>
        ///
        ///     byte[] b = new byte[10000];
        ///     for (int i = 0; i &lt; 10000; i++)
        ///     {
        ///         b[i % 120] = (byte)i;
        ///     }
        ///
        ///     CompoundFile cf = new CompoundFile();
        ///     CFStream myStream = cf.RootStorage.AddStream("MyStream");
        ///     Assert.IsNotNull(myStream);
        ///     myStream.SetData(b);
        ///     cf.Save("MyCompoundFile.cfs");
        ///     cf.Close();
        ///
        /// </code>
        /// </example>
        public CompoundFile()
        {
            _header = new Header();
            _sectorRecycle = false;

            _sectors.OnVer3SizeLimitReached += OnSizeLimitReached;

            _difatSectorFATEntriesCount = (GetSectorSize() / 4) - 1;
            _fatSectorEntriesCount = (GetSectorSize() / 4);

            //Root --
            var de = DirectoryEntry.New("Root Entry", StgType.StgRoot, _directoryEntries);
            RootStorage = new CFStorage(this, de)
            {
                DirEntry =
                {
                    StgType = StgType.StgRoot,
                    StgColor = StgColor.Black
                }
            };
        }

        /// <summary>
        /// Create a new, blank, compound file.
        /// </summary>
        /// <param name="cfsVersion">Use a specific Compound File Version to set 512 or 4096 bytes sectors</param>
        /// <param name="configFlags">Set <see cref="T:OpenMcdf.CFSConfiguration">configuration</see> parameters for the new compound file</param>
        /// <example>
        /// <code>
        ///
        ///     byte[] b = new byte[10000];
        ///     for (int i = 0; i &lt; 10000; i++)
        ///     {
        ///         b[i % 120] = (byte)i;
        ///     }
        ///
        ///     CompoundFile cf = new CompoundFile(CFSVersion.Ver_4, CFSConfiguration.Default);
        ///     CFStream myStream = cf.RootStorage.AddStream("MyStream");
        ///
        ///     Assert.IsNotNull(myStream);
        ///     myStream.SetData(b);
        ///     cf.Save("MyCompoundFile.cfs");
        ///     cf.Close();
        ///
        /// </code>
        /// </example>
        public CompoundFile(CFSVersion cfsVersion, CFSConfiguration configFlags)
        {
            Configuration = configFlags;

            var sectorRecycle = configFlags.HasFlag(CFSConfiguration.SectorRecycle);
            var eraseFreeSectors = configFlags.HasFlag(CFSConfiguration.EraseFreeSectors);

            _header = new Header((ushort)cfsVersion);
            _sectorRecycle = sectorRecycle;

            _difatSectorFATEntriesCount = (GetSectorSize() / 4) - 1;
            _fatSectorEntriesCount = (GetSectorSize() / 4);

            //Root --
            var rootDir = DirectoryEntry.New("Root Entry", StgType.StgRoot, _directoryEntries);
            rootDir.StgColor = StgColor.Black;
            RootStorage = new CFStorage(this, rootDir);
        }

        /// <summary>
        /// Load an existing compound file.
        /// </summary>
        /// <param name="fileName">Compound file to read from</param>
        /// <example>
        /// <code>
        /// //A xls file should have a Workbook stream
        /// string filename = "report.xls";
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
        /// <remarks>
        /// File will be open in read-only mode: it has to be saved
        /// with a different filename. A wrapping implementation has to be provided
        /// in order to remove/substitute an existing file. Version will be
        /// automatically recognized from the file. Sector recycle is turned off
        /// to achieve the best reading/writing performance in most common scenarios.
        /// </remarks>
        public CompoundFile(string fileName)
        {
            _sectorRecycle = false;
            UpdateMode = CFSUpdateMode.ReadOnly;
            _eraseFreeSectors = false;

            LoadFile(fileName);

            _difatSectorFATEntriesCount = (GetSectorSize() / 4) - 1;
            _fatSectorEntriesCount = (GetSectorSize() / 4);
        }

        public CompoundFile(MemoryStream memoryStream)
        {
            _sectorRecycle = false;
            UpdateMode = CFSUpdateMode.ReadOnly;
            _eraseFreeSectors = false;

            LoadStream(memoryStream);

            _difatSectorFATEntriesCount = (GetSectorSize() / 4) - 1;
            _fatSectorEntriesCount = (GetSectorSize() / 4);
        }

        ///  <summary>
        ///  Load an existing compound file.
        ///  </summary>
        ///  <param name="fileName">Compound file to read from</param>
        ///  <param name="updateMode">Select the update mode of the underlying data file</param>
        /// <param name="configParameters"></param>
        /// <example>
        ///  <code>
        ///  string srcFilename = "data_YOU_CAN_CHANGE.xls";
        ///
        ///  CompoundFile cf = new CompoundFile(srcFilename, UpdateMode.Update, true, true);
        ///
        ///  Random r = new Random();
        ///
        ///  byte[] buffer = GetBuffer(r.Next(3, 4095), 0x0A);
        ///
        ///  cf.RootStorage.AddStream("MyStream").SetData(buffer);
        ///
        ///  //This will persist data to the underlying media.
        ///  cf.Commit();
        ///  cf.Close();
        ///
        ///  </code>
        ///  </example>
        public CompoundFile(string fileName, CFSUpdateMode updateMode, CFSConfiguration configParameters)
        {
            ValidationExceptionEnabled = !configParameters.HasFlag(CFSConfiguration.NoValidationException);
            _sectorRecycle = configParameters.HasFlag(CFSConfiguration.SectorRecycle);
            UpdateMode = updateMode;
            _eraseFreeSectors = configParameters.HasFlag(CFSConfiguration.EraseFreeSectors);

            LoadFile(fileName);

            _difatSectorFATEntriesCount = (GetSectorSize() / 4) - 1;
            _fatSectorEntriesCount = (GetSectorSize() / 4);
        }

        ///  <summary>
        ///  Load an existing compound file.
        ///  </summary>
        ///  <param name="stream">A stream containing a compound file to read</param>
        ///  <param name="updateMode">Select the update mode of the underlying data file</param>
        /// <param name="configParameters"></param>
        /// <example>
        ///  <code>
        ///
        ///  string filename = "reportREAD.xls";
        ///
        ///  FileStream fs = new FileStream(filename, FileMode.Open);
        ///  CompoundFile cf = new CompoundFile(fs, UpdateMode.ReadOnly, false, false);
        ///  CFStream foundStream = cf.RootStorage.GetStream("Workbook");
        ///
        ///  byte[] temp = foundStream.GetData();
        ///
        ///  Assert.IsNotNull(temp);
        ///
        ///  cf.Close();
        ///
        ///  </code>
        ///  </example>
        ///  <exception cref="T:OpenMcdf.CFException">Raised when trying to open a non-seekable stream</exception>
        ///  <exception cref="T:OpenMcdf.CFException">Raised stream is null</exception>
        public CompoundFile(Stream stream, CFSUpdateMode updateMode, CFSConfiguration configParameters)
        {
            ValidationExceptionEnabled = !configParameters.HasFlag(CFSConfiguration.NoValidationException);
            _sectorRecycle = configParameters.HasFlag(CFSConfiguration.SectorRecycle);
            _eraseFreeSectors = configParameters.HasFlag(CFSConfiguration.EraseFreeSectors);
            _closeStream = !configParameters.HasFlag(CFSConfiguration.LeaveOpen);

            UpdateMode = updateMode;
            LoadStream(stream);

            _difatSectorFATEntriesCount = (GetSectorSize() / 4) - 1;
            _fatSectorEntriesCount = (GetSectorSize() / 4);
        }

        /// <inheritdoc />
        ///  <summary>
        ///  Load an existing compound file from a stream.
        ///  </summary>
        ///  <param name="stream">Streamed compound file</param>
        ///  <example>
        ///  <code>
        ///  string filename = "reportREAD.xls";
        ///  FileStream fs = new FileStream(filename, FileMode.Open);
        ///  CompoundFile cf = new CompoundFile(fs);
        ///  CFStream foundStream = cf.RootStorage.GetStream("Workbook");
        ///  byte[] temp = foundStream.GetData();
        ///  Assert.IsNotNull(temp);
        ///  cf.Close();
        ///  </code>
        ///  </example>
        ///  <exception cref="T:OpenMcdf.CFException">Raised when trying to open a non-seekable stream</exception>
        ///  <exception cref="T:OpenMcdf.CFException">Raised stream is null</exception>
        public CompoundFile(Stream stream) : this(stream, CFSUpdateMode.ReadOnly, CFSConfiguration.Default)
        {
        }

        /// <summary>
        /// Returns the size of standard sectors switching on CFS version (3 or 4)
        /// </summary>
        /// <returns>Standard sector size</returns>
        internal int GetSectorSize()
        {
            return 2 << (_header.SectorShift - 1);
        }

        private void OnSizeLimitReached()
        {
            using (var rangeLockSector = new Sector(GetSectorSize(), SourceStream))
            {
                _sectors.Add(rangeLockSector);

                rangeLockSector.Type = SectorType.RangeLockSector;

                TransactionLockAdded = true;
                LockSectorId = rangeLockSector.Id;
            }
        }

        /// <summary>
        /// Commit data changes since the previously commit operation
        /// to the underlying supporting stream or file on the disk.
        /// </summary>
        /// <remarks>
        /// This method can be used
        /// only if the supporting stream has been opened in
        /// <see cref="T:OpenMcdf.UpdateMode">Update mode</see>.
        /// </remarks>
        public void Commit()
        {
            Commit(false);
        }

#if !FLAT_WRITE
        private byte[] buffer = new byte[FLUSHING_BUFFER_MAX_SIZE];
        private Queue<Sector> flushingQueue = new Queue<Sector>(FLUSHING_QUEUE_SIZE);
#endif

        /// <summary>
        /// Commit data changes since the previously commit operation
        /// to the underlying supporting stream or file on the disk.
        /// </summary>
        /// <param name="releaseMemory">If true, release loaded sectors to limit memory usage but reduces following read operations performance</param>
        /// <remarks>
        /// This method can be used only if
        /// the supporting stream has been opened in
        /// <see cref="T:OpenMcdf.UpdateMode">Update mode</see>.
        /// </remarks>
        public void Commit(bool releaseMemory)
        {
            if (IsClosed)
                throw new CFDisposedException("Compound File closed: cannot commit data");

            if (UpdateMode != CFSUpdateMode.Update)
                throw new CFInvalidOperation("Cannot commit data in Read-Only update mode");

#if !FLAT_WRITE

            int sId = -1;
            int sCount = 0;
            int bufOffset = 0;
#endif
            var sSize = GetSectorSize();

            if (_header.MajorVersion != (ushort)CFSVersion.Ver_3)
                CheckForLockSector();

            SourceStream.Seek(0, SeekOrigin.Begin);
            SourceStream.Write(new byte[GetSectorSize()], 0, sSize);

            CommitDirectory();

            var gap = true;

            for (var i = 0; i < _sectors.Count; i++)
            {
#if FLAT_WRITE

                //Note:
                //Here sectors should not be loaded dynamically because
                //if they are null it means that no change has involved them;

                var s = _sectors[i];

                if (s != null && s.DirtyFlag)
                {
                    if (gap)
                        SourceStream.Seek(sSize + i * (long)sSize, SeekOrigin.Begin);

                    SourceStream.Write(s.GetData(), 0, sSize);
                    SourceStream.Flush();
                    s.DirtyFlag = false;
                    gap = false;
                }
                else
                {
                    gap = true;
                }

                if (!releaseMemory)
                    continue;

                s?.ReleaseData();
                _sectors[i] = null;

#else

                Sector s = sectors[i] as Sector;

                if (s != null && s.DirtyFlag && flushingQueue.Count < (int)(buffer.Length / sSize))
                {
                    //First of a block of contiguous sectors, mark id, start enqueuing

                    if (gap)
                    {
                        sId = s.Id;
                        gap = false;
                    }

                    flushingQueue.Enqueue(s);
                }
                else
                {
                    //Found a gap, stop enqueuing, flush a write operation

                    gap = true;
                    sCount = flushingQueue.Count;

                    if (sCount == 0) continue;

                    bufOffset = 0;
                    while (flushingQueue.Count > 0)
                    {
                        Sector r = flushingQueue.Dequeue();
                        Buffer.BlockCopy(r.GetData(), 0, buffer, bufOffset, sSize);
                        r.DirtyFlag = false;

                        if (releaseMemory)
                        {
                            r.ReleaseData();
                        }

                        bufOffset += sSize;
                    }

                    sourceStream.Seek(((long)sSize + (long)sId * (long)sSize), SeekOrigin.Begin);
                    sourceStream.Write(buffer, 0, sCount * sSize);

                    //Console.WriteLine("W - " + (int)(sCount * sSize ));
                }
#endif
            }

#if !FLAT_WRITE
            sCount = flushingQueue.Count;
            bufOffset = 0;

            while (flushingQueue.Count > 0)
            {
                Sector r = flushingQueue.Dequeue();
                Buffer.BlockCopy(r.GetData(), 0, buffer, bufOffset, sSize);
                r.DirtyFlag = false;

                if (releaseMemory)
                {
                    r.ReleaseData();
                    r = null;
                }

                bufOffset += sSize;
            }

            if (sCount != 0)
            {
                sourceStream.Seek((long)sSize + (long)sId * (long)sSize, SeekOrigin.Begin);
                sourceStream.Write(buffer, 0, sCount * sSize);
                //Console.WriteLine("W - " + (int)(sCount * sSize));
            }

#endif

            // Seek to beginning position and save header (first 512 or 4096 bytes)
            SourceStream.Seek(0, SeekOrigin.Begin);
            _header.Write(SourceStream);

            SourceStream.SetLength((_sectors.Count + 1) * sSize);
            SourceStream.Flush();

            if (releaseMemory)
                GC.Collect();
        }

        /// <summary>
        /// Load compound file from an existing stream.
        /// </summary>
        /// <param name="stream">Stream to load compound file from</param>
        private void Load(Stream stream)
        {
            try
            {
                _header = new Header();
                _directoryEntries = new List<IDirectoryEntry>();

                SourceStream = stream;

                _header.Read(stream);

                var nSector = Ceiling(((stream.Length - GetSectorSize()) / (double)GetSectorSize()));

                if (stream.Length > 0x7FFFFF0)
                    TransactionLockAllocated = true;

                _sectors = new SectorCollection();
                //sectors = new ArrayList();
                for (var i = 0; i < nSector; i++)
                {
                    _sectors.Add(null);
                }

                LoadDirectories();

                RootStorage
                    = new CFStorage(this, _directoryEntries[0]);
            }
            catch (Exception)
            {
                if (stream != null && _closeStream)
                    stream.Close();

                throw;
            }
        }

        private void LoadFile(string fileName)
        {
            FileName = fileName;

            FileStream fs = null;

            try
            {
                if (UpdateMode == CFSUpdateMode.ReadOnly)
                {
                    fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                }
                else
                {
                    fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
                }

                Load(fs);
            }
            catch
            {
                if (fs != null)
                    fs.Close();

                throw;
            }
        }

        private void LoadStream(Stream stream)
        {
            if (stream == null)
                throw new CFException("Stream parameter cannot be null");

            if (!stream.CanSeek)
                throw new CFException("Cannot load a non-seekable Stream");

            stream.Seek(0, SeekOrigin.Begin);

            Load(stream);
        }

        /// <summary>
        /// Return true if this compound file has been
        /// loaded from an existing file or stream
        /// </summary>
        public bool HasSourceStream => SourceStream != null;

        private void PersistMiniStreamToStream(IList<Sector> miniSectorChain)
        {
            var miniStream
                = GetSectorChain(RootEntry.StartSetc, SectorType.Normal);

            var miniStreamView
                = new StreamView(
                    miniStream,
                    GetSectorSize(),
                    RootStorage.Size,
                    null,
                    SourceStream);

            for (var i = 0; i < miniSectorChain.Count; i++)
            {
                var s = miniSectorChain[i];

                if (s.Id == -1)
                    throw new CFException("Invalid minisector index");

                // Ministream sectors already allocated
                miniStreamView.Seek(Sector.MINISECTOR_SIZE * s.Id, SeekOrigin.Begin);
                miniStreamView.Write(s.GetData(), 0, Sector.MINISECTOR_SIZE);
            }
        }

        /// <summary>
        /// Allocate space, setup sectors id and refresh header
        /// for the new or updated mini sector chain.
        /// </summary>
        /// <param name="sectorChain">The new MINI sector chain</param>
        private void AllocateMiniSectorChain(IList<Sector> sectorChain)
        {
            var miniFAT
                = GetSectorChain(_header.FirstMiniFATSectorId, SectorType.Normal);

            var miniStream
                = GetSectorChain(RootEntry.StartSetc, SectorType.Normal);

            var miniFATView
                = new StreamView(
                    miniFAT,
                    GetSectorSize(),
                    _header.MiniFATSectorsNumber * Sector.MINISECTOR_SIZE,
                    null,
                    SourceStream,
                    true
                );

            var miniStreamView
                = new StreamView(
                    miniStream,
                    GetSectorSize(),
                    RootStorage.Size,
                    null,
                    SourceStream);

            // Set updated/new sectors within the ministream
            // We are writing data in a NORMAL Sector chain.
            for (var i = 0; i < sectorChain.Count; i++)
            {
                var s = sectorChain[i];

                if (s.Id == -1)
                {
                    // Allocate, position ministream at the end of already allocated
                    // ministream's sectors

                    miniStreamView.Seek(RootStorage.Size + Sector.MINISECTOR_SIZE, SeekOrigin.Begin);
                    //miniStreamView.Write(s.GetData(), 0, Sector.MINISECTOR_SIZE);
                    s.Id = (int)(miniStreamView.Position - Sector.MINISECTOR_SIZE) / Sector.MINISECTOR_SIZE;

                    RootStorage.DirEntry.Size = miniStreamView.Length;
                }
            }

            // Update miniFAT
            for (var i = 0; i < sectorChain.Count - 1; i++)
            {
                var currentId = sectorChain[i].Id;
                var nextId = sectorChain[i + 1].Id;

                miniFATView.Seek(currentId * 4, SeekOrigin.Begin);
                miniFATView.Write(BitConverter.GetBytes(nextId), 0, 4);
            }

            // Write End of Chain in MiniFAT
            miniFATView.Seek(sectorChain[sectorChain.Count - 1].Id * SIZE_OF_SID, SeekOrigin.Begin);
            miniFATView.Write(BitConverter.GetBytes(Sector.ENDOFCHAIN), 0, 4);

            // Update sector chains
            AllocateSectorChain(miniStreamView.BaseSectorChain);
            AllocateSectorChain(miniFATView.BaseSectorChain);

            //Update HEADER and root storage when ministream changes
            if (miniFAT.Count > 0)
            {
                RootStorage.DirEntry.StartSetc = miniStream[0].Id;
                _header.MiniFATSectorsNumber = (uint)miniFAT.Count;
                _header.FirstMiniFATSectorId = miniFAT[0].Id;
            }
        }

        internal void FreeData(CFStream stream)
        {
            if (stream.Size == 0)
                return;

            List<Sector> sectorChain;

            if (stream.Size < _header.MinSizeStandardStream)
            {
                sectorChain = GetSectorChain(stream.DirEntry.StartSetc, SectorType.Mini);
                FreeMiniChain(sectorChain, _eraseFreeSectors);
            }
            else
            {
                sectorChain = GetSectorChain(stream.DirEntry.StartSetc, SectorType.Normal);
                FreeChain(sectorChain, _eraseFreeSectors);
            }

            stream.DirEntry.StartSetc = Sector.ENDOFCHAIN;
            stream.DirEntry.Size = 0;
        }

        private void FreeChain(IList<Sector> sectorChain, bool zeroSector)
        {
            FreeChain(sectorChain, 0, zeroSector);
        }

        private void FreeChain(IList<Sector> sectorChain, int nthSectorToRemove, bool zeroSector)
        {
            // Dummy zero buffer
            // ReSharper disable once InconsistentNaming
            var ZEROED_SECTOR = new byte[GetSectorSize()];

            var fat
                = GetSectorChain(-1, SectorType.FAT);

            var fatView
                = new StreamView(fat, GetSectorSize(), fat.Count * GetSectorSize(), null, SourceStream);

            // Zeroes out sector data (if required)-------------
            if (zeroSector)
            {
                for (var i = nthSectorToRemove; i < sectorChain.Count; i++)
                {
                    var s = sectorChain[i];
                    s.ZeroData();
                }
            }

            // Update FAT marking unallocated sectors ----------
            for (var i = nthSectorToRemove; i < sectorChain.Count; i++)
            {
                var currentId = sectorChain[i].Id;

                fatView.Seek(currentId * 4, SeekOrigin.Begin);
                fatView.Write(BitConverter.GetBytes(Sector.FREESECT), 0, 4);
            }

            // Write new end of chain if partial free ----------
            if (nthSectorToRemove > 0 && sectorChain.Count > 0)
            {
                fatView.Seek(sectorChain[nthSectorToRemove - 1].Id * 4, SeekOrigin.Begin);
                fatView.Write(BitConverter.GetBytes(Sector.ENDOFCHAIN), 0, 4);
            }
        }

        private void FreeMiniChain(IList<Sector> sectorChain, bool zeroSector)
        {
            FreeMiniChain(sectorChain, 0, zeroSector);
        }

        private void FreeMiniChain(IList<Sector> sectorChain, int nthSectorToRemove, bool zeroSector)
        {
            var zeroedMiniSector = new byte[Sector.MINISECTOR_SIZE];

            var miniFAT
                = GetSectorChain(_header.FirstMiniFATSectorId, SectorType.Normal);

            var miniStream
                = GetSectorChain(RootEntry.StartSetc, SectorType.Normal);

            var miniFATView
                = new StreamView(miniFAT, GetSectorSize(), _header.MiniFATSectorsNumber * Sector.MINISECTOR_SIZE, null,
                    SourceStream);

            var miniStreamView
                = new StreamView(miniStream, GetSectorSize(), RootStorage.Size, null, SourceStream);

            // Set updated/new sectors within the mini-stream ----------
            if (zeroSector)
            {
                for (var i = nthSectorToRemove; i < sectorChain.Count; i++)
                {
                    var s = sectorChain[i];

                    if (s.Id == -1)
                        continue;

                    // Overwrite
                    miniStreamView.Seek(Sector.MINISECTOR_SIZE * s.Id, SeekOrigin.Begin);
                    miniStreamView.Write(zeroedMiniSector, 0, Sector.MINISECTOR_SIZE);
                }
            }

            // Update miniFAT                ---------------------------------------
            for (var i = nthSectorToRemove; i < sectorChain.Count; i++)
            {
                var currentId = sectorChain[i].Id;

                miniFATView.Seek(currentId * 4, SeekOrigin.Begin);
                miniFATView.Write(BitConverter.GetBytes(Sector.FREESECT), 0, 4);
            }

            // Write End of Chain in MiniFAT ---------------------------------------
            //miniFATView.Seek(sectorChain[(sectorChain.Count - 1) - nth_sector_to_remove].Id * SIZE_OF_SID, SeekOrigin.Begin);
            //miniFATView.Write(BitConverter.GetBytes(Sector.ENDOFCHAIN), 0, 4);

            // Write End of Chain in MiniFAT ---------------------------------------
            if (nthSectorToRemove > 0 && sectorChain.Count > 0)
            {
                miniFATView.Seek(sectorChain[nthSectorToRemove - 1].Id * 4, SeekOrigin.Begin);
                miniFATView.Write(BitConverter.GetBytes(Sector.ENDOFCHAIN), 0, 4);
            }

            // Update sector chains           ---------------------------------------
            AllocateSectorChain(miniStreamView.BaseSectorChain);
            AllocateSectorChain(miniFATView.BaseSectorChain);

            //Update HEADER and root storage when mini-stream changes
            if (miniFAT.Count > 0)
            {
                RootStorage.DirEntry.StartSetc = miniStream[0].Id;
                _header.MiniFATSectorsNumber = (uint)miniFAT.Count;
                _header.FirstMiniFATSectorId = miniFAT[0].Id;
            }
        }

        /// <summary>
        /// Allocate space, setup sectors id in the FAT and refresh header
        /// for the new or updated sector chain (Normal or Mini sectors)
        /// </summary>
        /// <param name="sectorChain">The new or updated normal or mini sector chain</param>
        private void SetSectorChain(IList<Sector> sectorChain)
        {
            if (sectorChain == null || sectorChain.Count == 0)
                return;

            var st = sectorChain[0].Type;

            if (st == SectorType.Normal)
            {
                AllocateSectorChain(sectorChain);
            }
            else if (st == SectorType.Mini)
            {
                AllocateMiniSectorChain(sectorChain);
            }
        }

        /// <summary>
        /// Allocate space, setup sectors id and refresh header
        /// for the new or updated sector chain.
        /// </summary>
        /// <param name="sectorChain">The new or updated generic sector chain</param>
        private void AllocateSectorChain(IList<Sector> sectorChain)
        {
            foreach (var s in sectorChain)
            {
                if (s.Id == -1)
                {
                    _sectors.Add(s);
                    s.Id = _sectors.Count - 1;
                }
            }

            AllocateFATSectorChain(sectorChain);
        }

        /// <summary>
        /// Check for transaction lock sector addition and mark it in the FAT.
        /// </summary>
        private void CheckForLockSector()
        {
            //If transaction lock has been added and not yet allocated in the FAT...
            if (!TransactionLockAdded || TransactionLockAllocated)
                return;

            using (var fatStream = new StreamView(GetFatSectorChain(), GetSectorSize(), SourceStream))
            {
                fatStream.Seek(LockSectorId * 4, SeekOrigin.Begin);
                fatStream.Write(BitConverter.GetBytes(Sector.ENDOFCHAIN), 0, 4);

                TransactionLockAllocated = true;
            }
        }

        /// <summary>
        /// Allocate space, setup sectors id and refresh header
        /// for the new or updated FAT sector chain.
        /// </summary>
        /// <param name="sectorChain">The new or updated generic sector chain</param>
        private void AllocateFATSectorChain(IList<Sector> sectorChain)
        {
            var fatSectors = GetSectorChain(-1, SectorType.FAT);

            var fatStream =
                new StreamView(
                    fatSectors,
                    GetSectorSize(),
                    _header.FATSectorsNumber * GetSectorSize(),
                    null,
                    SourceStream,
                    true
                );

            // Write FAT chain values --

            for (var i = 0; i < sectorChain.Count - 1; i++)
            {
                var sN = sectorChain[i + 1];
                var sC = sectorChain[i];

                fatStream.Seek(sC.Id * 4, SeekOrigin.Begin);
                fatStream.Write(BitConverter.GetBytes(sN.Id), 0, 4);
            }

            fatStream.Seek(sectorChain[sectorChain.Count - 1].Id * 4, SeekOrigin.Begin);
            fatStream.Write(BitConverter.GetBytes(Sector.ENDOFCHAIN), 0, 4);

            // Merge chain to CFS
            AllocateDIFATSectorChain(fatStream.BaseSectorChain);
        }

        /// <summary>
        /// Setup the DIFAT sector chain
        /// </summary>
        /// <param name="faTsectorChain">A FAT sector chain</param>
        private void AllocateDIFATSectorChain(IList<Sector> faTsectorChain)
        {
            // Get initial sector's count
            _header.FATSectorsNumber = faTsectorChain.Count;

            // Allocate Sectors
            foreach (var s in faTsectorChain)
            {
                if (s.Id != -1)
                    continue;

                _sectors.Add(s);
                s.Id = _sectors.Count - 1;
                s.Type = SectorType.FAT;
            }

            // Sector count...
            var nCurrentSectors = _sectors.Count;

            // Temp DIFAT count
            var nDIFATSectors = (int)_header.DIFATSectorsNumber;

            if (faTsectorChain.Count > HEADER_DIFAT_ENTRIES_COUNT)
            {
                nDIFATSectors = Ceiling((double)(faTsectorChain.Count - HEADER_DIFAT_ENTRIES_COUNT) /
                                        _difatSectorFATEntriesCount);
                nDIFATSectors = LowSaturation(nDIFATSectors - (int)_header.DIFATSectorsNumber); //required DIFAT
            }

            // ...sum with new required DIFAT sectors count
            nCurrentSectors += nDIFATSectors;

            // ReCheck FAT bias
            while (_header.FATSectorsNumber * _fatSectorEntriesCount < nCurrentSectors)
            {
                var extraFATSector = new Sector(GetSectorSize(), SourceStream);
                _sectors.Add(extraFATSector);

                extraFATSector.Id = _sectors.Count - 1;
                extraFATSector.Type = SectorType.FAT;

                faTsectorChain.Add(extraFATSector);

                _header.FATSectorsNumber++;
                nCurrentSectors++;

                //... so, adding a FAT sector may induce DIFAT sectors to increase by one
                // and consequently this may induce ANOTHER FAT sector (TO-THINK: May this condition occure ?)
                if (nDIFATSectors * _difatSectorFATEntriesCount >=
                    (_header.FATSectorsNumber > HEADER_DIFAT_ENTRIES_COUNT
                        ? _header.FATSectorsNumber - HEADER_DIFAT_ENTRIES_COUNT
                        : 0))
                {
                    continue;
                }

                nDIFATSectors++;
                nCurrentSectors++;
            }

            var difatSectors =
                GetSectorChain(-1, SectorType.DIFAT);

            var difatStream
                = new StreamView(difatSectors, GetSectorSize(), SourceStream);

            // Write DIFAT Sectors (if required)
            // Save room for the following chaining
            for (var i = 0; i < faTsectorChain.Count; i++)
            {
                if (i < HEADER_DIFAT_ENTRIES_COUNT)
                {
                    _header.DIFAT[i] = faTsectorChain[i].Id;
                }
                else
                {
                    // room for DIFAT chaining at the end of any DIFAT sector (4 bytes)
                    if (i != HEADER_DIFAT_ENTRIES_COUNT &&
                        (i - HEADER_DIFAT_ENTRIES_COUNT) % _difatSectorFATEntriesCount == 0)
                    {
                        var temp = new byte[sizeof(int)];
                        difatStream.Write(temp, 0, sizeof(int));
                    }

                    difatStream.Write(BitConverter.GetBytes(faTsectorChain[i].Id), 0, sizeof(int));
                }
            }

            // Allocate room for DIFAT sectors
            foreach (var sec in difatStream.BaseSectorChain)
            {
                if (sec.Id != -1)
                    continue;

                _sectors.Add(sec);
                sec.Id = _sectors.Count - 1;
                sec.Type = SectorType.DIFAT;
            }

            _header.DIFATSectorsNumber = (uint)nDIFATSectors;

            // Chain first sector
            if (difatStream.BaseSectorChain != null && difatStream.BaseSectorChain.Count > 0)
            {
                _header.FirstDIFATSectorId = difatStream.BaseSectorChain[0].Id;

                // Update header information
                _header.DIFATSectorsNumber = (uint)difatStream.BaseSectorChain.Count;

                // Write chaining information at the end of DIFAT Sectors
                for (var i = 0; i < difatStream.BaseSectorChain.Count - 1; i++)
                {
                    Buffer.BlockCopy(
                        BitConverter.GetBytes(difatStream.BaseSectorChain[i + 1].Id),
                        0,
                        difatStream.BaseSectorChain[i].GetData(),
                        GetSectorSize() - sizeof(int),
                        4);
                }

                Buffer.BlockCopy(
                    BitConverter.GetBytes(Sector.ENDOFCHAIN),
                    0,
                    difatStream.BaseSectorChain[difatStream.BaseSectorChain.Count - 1].GetData(),
                    GetSectorSize() - sizeof(int),
                    sizeof(int)
                );
            }
            else
                _header.FirstDIFATSectorId = Sector.ENDOFCHAIN;

            // Mark DIFAT Sectors in FAT
            var fatSv =
                new StreamView(faTsectorChain, GetSectorSize(), _header.FATSectorsNumber * GetSectorSize(), null,
                    SourceStream);

            for (var i = 0; i < _header.DIFATSectorsNumber; i++)
            {
                fatSv.Seek(difatStream.BaseSectorChain[i].Id * 4, SeekOrigin.Begin);
                fatSv.Write(BitConverter.GetBytes(Sector.DIFSECT), 0, 4);
            }

            for (var i = 0; i < _header.FATSectorsNumber; i++)
            {
                fatSv.Seek(fatSv.BaseSectorChain[i].Id * 4, SeekOrigin.Begin);
                fatSv.Write(BitConverter.GetBytes(Sector.FATSECT), 0, 4);
            }

            //fatSv.Seek(fatSv.BaseSectorChain[fatSv.BaseSectorChain.Count - 1].Id * 4, SeekOrigin.Begin);
            //fatSv.Write(BitConverter.GetBytes(Sector.ENDOFCHAIN), 0, 4);

            _header.FATSectorsNumber = fatSv.BaseSectorChain.Count;
        }

        /// <summary>
        /// Get the DIFAT Sector chain
        /// </summary>
        /// <returns>A list of DIFAT sectors</returns>
        private List<Sector> GetDifatSectorChain()
        {
            int validationCount;

            var result
                = new List<Sector>();

            if (_header.DIFATSectorsNumber == 0)
                return result;

            validationCount = (int)_header.DIFATSectorsNumber;

            if (!(_sectors[_header.FirstDIFATSectorId] is Sector s)) //Lazy loading
            {
                s = new Sector(GetSectorSize(), SourceStream)
                {
                    Type = SectorType.DIFAT,
                    Id = _header.FirstDIFATSectorId
                };
                _sectors[_header.FirstDIFATSectorId] = s;
            }

            result.Add(s);

            while (validationCount >= 0)
            {
                var nextSecId = BitConverter.ToInt32(s.GetData(), GetSectorSize() - 4);

                // Strictly speaking, the following condition is not correct from
                // a specification point of view:
                // only ENDOFCHAIN should break DIFAT chain but
                // a lot of existing compound files use FREESECT as DIFAT chain termination
                if (nextSecId == Sector.FREESECT || nextSecId == Sector.ENDOFCHAIN) break;

                validationCount--;

                if (validationCount < 0)
                {
                    Close();
                    throw new CFCorruptedFileException("DIFAT sectors count mismatched. Corrupted compound file");
                }

                s = _sectors[nextSecId];

                if (s == null)
                {
                    s = new Sector(GetSectorSize(), SourceStream);
                    s.Id = nextSecId;
                    _sectors[nextSecId] = s;
                }

                result.Add(s);
            }

            return result;
        }

        /// <summary>
        /// Get the FAT sector chain
        /// </summary>
        /// <returns>List of FAT sectors</returns>
        private List<Sector> GetFatSectorChain()
        {
            const int nHeaderFATEntry = 109; //Number of FAT sectors id in the header

            var result
                = new List<Sector>();

            int nextSecId;

            var difatSectors = GetDifatSectorChain();

            var idx = 0;

            // Read FAT entries from the header Fat entry array (max 109 entries)
            while (idx < _header.FATSectorsNumber && idx < nHeaderFATEntry)
            {
                nextSecId = _header.DIFAT[idx];
                var s = _sectors[nextSecId];

                if (s == null)
                {
                    s = new Sector(GetSectorSize(), SourceStream);
                    s.Id = nextSecId;
                    s.Type = SectorType.FAT;
                    _sectors[nextSecId] = s;
                }

                result.Add(s);

                idx++;
            }

            //Is there any DIFAT sector containing other FAT entries ?
            if (difatSectors.Count > 0)
            {
                var difatStream
                    = new StreamView
                    (
                        difatSectors,
                        GetSectorSize(),
                        _header.FATSectorsNumber > nHeaderFATEntry
                            ? (_header.FATSectorsNumber - nHeaderFATEntry) * 4
                            : 0,
                        null,
                        SourceStream
                    );

                var nextDIFATSectorBuffer = new byte[4];

                difatStream.Read(nextDIFATSectorBuffer, 0, 4);
                nextSecId = BitConverter.ToInt32(nextDIFATSectorBuffer, 0);

                var i = 0;
                var nFat = nHeaderFATEntry;

                while (nFat < _header.FATSectorsNumber)
                {
                    if (difatStream.Position == ((GetSectorSize() - 4) + i * GetSectorSize()))
                    {
                        difatStream.Seek(4, SeekOrigin.Current);
                        i++;
                        continue;
                    }

                    var s = _sectors[nextSecId];

                    if (s == null)
                    {
                        s = new Sector(GetSectorSize(), SourceStream);
                        s.Type = SectorType.FAT;
                        s.Id = nextSecId;
                        _sectors[nextSecId] = s; //UUU
                    }

                    result.Add(s);

                    difatStream.Read(nextDIFATSectorBuffer, 0, 4);
                    nextSecId = BitConverter.ToInt32(nextDIFATSectorBuffer, 0);
                    nFat++;
                }
            }

            return result;
        }

        /// <summary>
        /// Get a standard sector chain
        /// </summary>
        /// <param name="secId">First SecID of the required chain</param>
        /// <returns>A list of sectors</returns>
        private List<Sector> GetNormalSectorChain(int secId)
        {
            var result
                = new List<Sector>();

            var nextSecId = secId;

            var fatSectors = GetFatSectorChain();

            var fatStream
                = new StreamView(fatSectors, GetSectorSize(), fatSectors.Count * GetSectorSize(), null, SourceStream);

            while (true)
            {
                if (nextSecId == Sector.ENDOFCHAIN) break;

                if (nextSecId < 0)
                    throw new CFCorruptedFileException(
                        string.Format("Next Sector ID reference is below zero. NextID : {0}", nextSecId));

                if (nextSecId >= _sectors.Count)
                    throw new CFCorruptedFileException(string.Format(
                        "Next Sector ID reference an out of range sector. NextID : {0} while sector count {1}",
                        nextSecId, _sectors.Count));

                if (!(_sectors[nextSecId] is Sector s))
                {
                    s = new Sector(GetSectorSize(), SourceStream);
                    s.Id = nextSecId;
                    s.Type = SectorType.Normal;
                    _sectors[nextSecId] = s;
                }

                result.Add(s);

                fatStream.Seek(nextSecId * 4, SeekOrigin.Begin);
                var next = fatStream.ReadInt32();

                if (next != nextSecId)
                    nextSecId = next;
                else
                    throw new CFCorruptedFileException("Cyclic sector chain found. File is corrupted");
            }

            return result;
        }

        /// <summary>
        /// Get a mini sector chain
        /// </summary>
        /// <param name="secId">First SecID of the required chain</param>
        /// <returns>A list of mini sectors (64 bytes)</returns>
        private List<Sector> GetMiniSectorChain(int secId)
        {
            var result
                = new List<Sector>();

            //if (secId != Sector.ENDOFCHAIN)
            {
                int nextSecId;

                var miniFAT = GetNormalSectorChain(_header.FirstMiniFATSectorId);
                var miniStream = GetNormalSectorChain(RootEntry.StartSetc);

                var miniFATView
                    = new StreamView(miniFAT, GetSectorSize(), _header.MiniFATSectorsNumber * Sector.MINISECTOR_SIZE,
                        null, SourceStream);

                var miniStreamView =
                    new StreamView(miniStream, GetSectorSize(), RootStorage.Size, null, SourceStream);

                var miniFATReader = new BinaryReader(miniFATView);

                nextSecId = secId;

                while (true)
                {
                    if (nextSecId == Sector.ENDOFCHAIN)
                        break;

                    var ms = new Sector(Sector.MINISECTOR_SIZE, SourceStream);
                    // ReSharper disable once UnusedVariable
                    var temp = new byte[Sector.MINISECTOR_SIZE];

                    ms.Id = nextSecId;
                    ms.Type = SectorType.Mini;

                    miniStreamView.Seek(nextSecId * Sector.MINISECTOR_SIZE, SeekOrigin.Begin);
                    miniStreamView.Read(ms.GetData(), 0, Sector.MINISECTOR_SIZE);

                    result.Add(ms);

                    miniFATView.Seek(nextSecId * 4, SeekOrigin.Begin);
                    nextSecId = miniFATReader.ReadInt32();
                }
            }
            return result;
        }

        /// <summary>
        /// Get a sector chain from a compound file given the first sector ID
        /// and the required sector type.
        /// </summary>
        /// <param name="secId">First chain sector's id </param>
        /// <param name="chainType">Type of Sectors in the required chain (mini sectors, normal sectors or FAT)</param>
        /// <returns>A list of Sectors as the result of their concatenation</returns>
        internal List<Sector> GetSectorChain(int secId, SectorType chainType)
        {
            switch (chainType)
            {
                case SectorType.DIFAT:
                    return GetDifatSectorChain();

                case SectorType.FAT:
                    return GetFatSectorChain();

                case SectorType.Normal:
                    return GetNormalSectorChain(secId);

                case SectorType.Mini:
                    return GetMiniSectorChain(secId);

                default:
                    throw new CFException("Unsupproted chain type");
            }
        }

        /// <summary>
        /// Reset a directory entry setting it to StgInvalid in the Directory.
        /// </summary>
        /// <param name="sid">Sid of the directory to invalidate</param>
        internal void ResetDirectoryEntry(int sid)
        {
            _directoryEntries[sid].SetEntryName(string.Empty);
            _directoryEntries[sid].Left = null;
            _directoryEntries[sid].Right = null;
            _directoryEntries[sid].Parent = null;
            _directoryEntries[sid].StgType = StgType.StgInvalid;
        }

        internal RbTree CreateNewTree()
        {
            var bst = new RbTree();
            //bst.NodeInserted += OnNodeInsert;
            //bst.NodeOperation += OnNodeOperation;
            //bst.NodeDeleted += new Action<RBNode<CFItem>>(OnNodeDeleted);
            //  bst.ValueAssignedAction += new Action<RBNode<CFItem>, CFItem>(OnValueAssigned);
            return bst;
        }

        internal RbTree GetChildrenTree(int sid)
        {
            var bst = new RbTree();

            // Load children from their original tree.
            DoLoadChildren(bst, _directoryEntries[sid]);

            return bst;
        }

        internal RbTree DoLoadChildrenTrusted(IDirectoryEntry de)
        {
            RbTree bst = null;

            if (de.Child != DirectoryEntry.nostream)
            {
                bst = new RbTree(_directoryEntries[de.Child]);
            }

            return bst;
        }

        private void DoLoadChildren(RbTree bst, IDirectoryEntry de)
        {
            if (de.Child != DirectoryEntry.nostream)
            {
                if (_directoryEntries[de.Child].StgType == StgType.StgInvalid) return;

                LoadSiblings(bst, _directoryEntries[de.Child]);
                NullifyChildNodes(_directoryEntries[de.Child]);
                bst.Insert(_directoryEntries[de.Child]);
            }
        }

        private static void NullifyChildNodes(IRbNode de)
        {
            de.Parent = null;
            de.Left = null;
            de.Right = null;
        }

        // Doubling methods allows iterative behavior while avoiding
        // to insert duplicate items
        private void LoadSiblings(RbTree bst, IDirectoryEntry de)
        {
            _levelSiDs.Clear();

            if (de.LeftSibling != DirectoryEntry.nostream)
            {
                // If there're more left siblings load them...
                DoLoadSiblings(bst, _directoryEntries[de.LeftSibling]);
                //NullifyChildNodes(directoryEntries[de.LeftSibling]);
            }

            if (de.RightSibling != DirectoryEntry.nostream)
            {
                _levelSiDs.Add(de.RightSibling);

                // If there're more right siblings load them...
                DoLoadSiblings(bst, _directoryEntries[de.RightSibling]);
                //NullifyChildNodes(directoryEntries[de.RightSibling]);
            }
        }

        private void DoLoadSiblings(RbTree bst, IDirectoryEntry de)
        {
            if (ValidateSibling(de.LeftSibling))
            {
                _levelSiDs.Add(de.LeftSibling);

                // If there're more left siblings load them...
                DoLoadSiblings(bst, _directoryEntries[de.LeftSibling]);
            }

            if (ValidateSibling(de.RightSibling))
            {
                _levelSiDs.Add(de.RightSibling);

                // If there're more right siblings load them...
                DoLoadSiblings(bst, _directoryEntries[de.RightSibling]);
            }

            NullifyChildNodes(de);
            bst.Insert(de);
        }

        private bool ValidateSibling(int sid)
        {
            if (sid != DirectoryEntry.nostream)
            {
                // if this siblings id does not overflow current list
                if (sid >= _directoryEntries.Count)
                {
                    if (ValidationExceptionEnabled)
                    {
                        //this.Close();
                        throw new CFCorruptedFileException("A Directory Entry references the non-existent sid number " +
                                                           sid);
                    }
                    return false;
                }

                //if this sibling is valid...
                if (_directoryEntries[sid].StgType == StgType.StgInvalid)
                {
                    if (ValidationExceptionEnabled)
                    {
                        //this.Close();
                        throw new CFCorruptedFileException(
                            "A Directory Entry has a valid reference to an Invalid Storage Type directory [" + sid +
                            "]");
                    }
                    return false;
                }

                if (!Enum.IsDefined(typeof(StgType), _directoryEntries[sid].StgType))
                {
                    if (ValidationExceptionEnabled)
                    {
                        //this.Close();
                        throw new CFCorruptedFileException("A Directory Entry has an invalid Storage Type");
                    }
                    return false;
                }

                if (_levelSiDs.Contains(sid))
                    throw new CFCorruptedFileException("Cyclic reference of directory item");

                return true; //No fault condition encountered for sid being validated
            }

            return false;
        }

        /// <summary>
        /// Load directory entries from compound file. Header and FAT MUST be already loaded.
        /// </summary>
        private void LoadDirectories()
        {
            var directoryChain
                = GetSectorChain(_header.FirstDirectorySectorId, SectorType.Normal);

            if (_header.FirstDirectorySectorId == Sector.ENDOFCHAIN)
                _header.FirstDirectorySectorId = directoryChain[0].Id;

            var dirReader
                = new StreamView(directoryChain, GetSectorSize(), directoryChain.Count * GetSectorSize(), null,
                    SourceStream);

            while (dirReader.Position < directoryChain.Count * GetSectorSize())
            {
                var de
                    = DirectoryEntry.New(string.Empty, StgType.StgInvalid, _directoryEntries);

                //We are not inserting dirs. Do not use 'InsertNewDirectoryEntry'
                de.Read(dirReader, Version);
            }
        }

        /// <summary>
        ///  Commit directory entries change on the Current Source stream
        /// </summary>
        private void CommitDirectory()
        {
            const int directorySize = 128;

            var directorySectors
                = GetSectorChain(_header.FirstDirectorySectorId, SectorType.Normal);

            var sv = new StreamView(directorySectors, GetSectorSize(), 0, null, SourceStream);

            foreach (var di in _directoryEntries)
            {
                di.Write(sv);
            }

            var delta = _directoryEntries.Count;

            while (delta % (GetSectorSize() / directorySize) != 0)
            {
                var dummy = DirectoryEntry.New(string.Empty, StgType.StgInvalid, _directoryEntries);
                dummy.Write(sv);
                delta++;
            }

            foreach (var s in directorySectors)
            {
                s.Type = SectorType.Directory;
            }

            AllocateSectorChain(directorySectors);

            _header.FirstDirectorySectorId = directorySectors[0].Id;

            //Version 4 supports directory sectors count
            if (_header.MajorVersion == 3)
            {
                _header.DirectorySectorsNumber = 0;
            }
            else
            {
                _header.DirectorySectorsNumber = directorySectors.Count;
            }
        }

        /// <summary>
        /// Saves the in-memory image of Compound File to a file.
        /// </summary>
        /// <param name="fileName">File name to write the compound file to</param>
        /// <exception cref="T:OpenMcdf.CFException">Raised if destination file is not seekable</exception>
        public void Save(string fileName)
        {
            if (IsClosed)
                throw new CFException("Compound File closed: cannot save data");

            FileStream fs = null;

            try
            {
                fs = new FileStream(fileName, FileMode.Create);
                Save(fs);
            }
            catch (Exception ex)
            {
                throw new CFException("Error saving file [" + fileName + "]", ex);
            }
            finally
            {
                fs?.Flush();
                fs?.Close();
            }
        }

        /// <summary>
        /// Saves the in-memory image of Compound File to a stream.
        /// </summary>
        /// <remarks>
        /// Destination Stream must be seekable. Uncommitted data will be persisted to the destination stream.
        /// </remarks>
        /// <param name="stream">The stream to save compound File to</param>
        /// <exception cref="T:OpenMcdf.CFException">Raised if destination stream is not seekable</exception>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised if Compound File Storage has been already disposed</exception>
        /// <example>
        /// <code>
        ///    MemoryStream ms = new MemoryStream(size);
        ///
        ///    CompoundFile cf = new CompoundFile();
        ///    CFStorage st = cf.RootStorage.AddStorage("MyStorage");
        ///    CFStream sm = st.AddStream("MyStream");
        ///
        ///    byte[] b = new byte[]{0x00,0x01,0x02,0x03};
        ///
        ///    sm.SetData(b);
        ///    cf.Save(ms);
        ///    cf.Close();
        /// </code>
        /// </example>
        public void Save(Stream stream)
        {
            if (IsClosed)
                throw new CFDisposedException("Compound File closed: cannot save data");

            if (!stream.CanSeek)
                throw new CFException("Cannot save on a non-seekable stream");

            CheckForLockSector();
            var sSize = GetSectorSize();

            try
            {
                stream.Write((byte[])Array.CreateInstance(typeof(byte), sSize), 0, sSize);

                CommitDirectory();

                for (var i = 0; i < _sectors.Count; i++)
                {
                    var s = _sectors[i];

                    if (s == null)
                    {
                        // Load source (unmodified) sectors
                        // Here we have to ignore "Dirty flag" of
                        // sectors because we are NOT modifying the source
                        // in a differential way but ALL sectors need to be
                        // persisted on the destination stream
                        s = new Sector(sSize, SourceStream) { Id = i };

                        //sectors[i] = s;
                    }

                    stream.Write(s.GetData(), 0, sSize);

                    //s.ReleaseData();
                }

                stream.Seek(0, SeekOrigin.Begin);
                _header.Write(stream);
            }
            catch (Exception ex)
            {
                throw new CFException("Internal error while saving compound file to stream ", ex);
            }
        }

        /// <summary>
        /// Scan FAT o miniFAT for free sectors to reuse.
        /// </summary>
        /// <param name="sType">Type of sector to look for</param>
        /// <returns>A Queue of available sectors or minisectors already allocated</returns>
        internal Queue<Sector> FindFreeSectors(SectorType sType)
        {
            var freeList = new Queue<Sector>();

            if (sType == SectorType.Normal)
            {
                var fatChain = GetSectorChain(-1, SectorType.FAT);
                var fatStream = new StreamView(fatChain, GetSectorSize(),
                    _header.FATSectorsNumber * GetSectorSize(), null, SourceStream);

                var idx = 0;

                while (idx < _sectors.Count)
                {
                    var id = fatStream.ReadInt32();

                    if (id == Sector.FREESECT)
                    {
                        if (_sectors[idx] == null)
                        {
                            var s = new Sector(GetSectorSize(), SourceStream);
                            s.Id = idx;
                            _sectors[idx] = s;
                        }

                        freeList.Enqueue(_sectors[idx]);
                    }

                    idx++;
                }
            }
            else
            {
                var miniFAT
                    = GetSectorChain(_header.FirstMiniFATSectorId, SectorType.Normal);

                var miniFATView
                    = new StreamView(miniFAT, GetSectorSize(), _header.MiniFATSectorsNumber * Sector.MINISECTOR_SIZE,
                        null, SourceStream);

                var miniStream
                    = GetSectorChain(RootEntry.StartSetc, SectorType.Normal);

                var miniStreamView
                    = new StreamView(miniStream, GetSectorSize(), RootStorage.Size, null, SourceStream);

                var idx = 0;

                var nMinisectors = (int)(miniStreamView.Length / Sector.MINISECTOR_SIZE);

                while (idx < nMinisectors)
                {
                    //AssureLength(miniStreamView, (int)miniFATView.Length);

                    var nextId = miniFATView.ReadInt32();

                    if (nextId == Sector.FREESECT)
                    {
                        var ms = new Sector(Sector.MINISECTOR_SIZE, SourceStream);
                        // ReSharper disable once UnusedVariable
                        var temp = new byte[Sector.MINISECTOR_SIZE];

                        ms.Id = idx;
                        ms.Type = SectorType.Mini;

                        miniStreamView.Seek(ms.Id * Sector.MINISECTOR_SIZE, SeekOrigin.Begin);
                        miniStreamView.Read(ms.GetData(), 0, Sector.MINISECTOR_SIZE);

                        freeList.Enqueue(ms);
                    }

                    idx++;
                }
            }

            return freeList;
        }

        /// <summary>
        /// INTERNAL DEVELOPMENT. DO NOT CALL.
        /// <param name="cfItem"></param>
        /// <param name="buffer"></param>
        /// </summary>
        internal void AppendData(CFItem cfItem, byte[] buffer)
        {
            WriteData(cfItem, cfItem.Size, buffer);
        }

        /// <summary>
        /// Resize stream length
        /// </summary>
        /// <param name="cfItem"></param>
        /// <param name="length"></param>
        internal void SetStreamLength(CFItem cfItem, long length)
        {
            if (cfItem.Size == length)
                return;

            var newSectorType = SectorType.Normal;
            var newSectorSize = GetSectorSize();

            if (length < _header.MinSizeStandardStream)
            {
                newSectorType = SectorType.Mini;
                newSectorSize = Sector.MINISECTOR_SIZE;
            }

            var oldSectorType = SectorType.Normal;
            var oldSectorSize = GetSectorSize();

            if (cfItem.Size < _header.MinSizeStandardStream)
            {
                oldSectorType = SectorType.Mini;
                oldSectorSize = Sector.MINISECTOR_SIZE;
            }

            var oldSize = cfItem.Size;

            // Get Sector chain and delta size induced by client
            var sectorChain = GetSectorChain(cfItem.DirEntry.StartSetc, oldSectorType);
            var delta = length - cfItem.Size;

            // Check for transition ministream -> stream:
            // Only in this case we need to free old sectors,
            // otherwise they will be overwritten.

            var transitionToMini = false;
            var transitionToNormal = false;
            List<Sector> oldChain = null;

            if (cfItem.DirEntry.StartSetc != Sector.ENDOFCHAIN)
            {
                if (
                    (length < _header.MinSizeStandardStream && cfItem.DirEntry.Size >= _header.MinSizeStandardStream)
                    || (length >= _header.MinSizeStandardStream && cfItem.DirEntry.Size < _header.MinSizeStandardStream)
                )
                {
                    if (cfItem.DirEntry.Size < _header.MinSizeStandardStream)
                    {
                        transitionToNormal = true;
                        oldChain = sectorChain;
                    }
                    else
                    {
                        transitionToMini = true;
                        oldChain = sectorChain;
                    }

                    // No transition caused by size change
                }
            }

            Queue<Sector> freeList = null;
            StreamView sv;

            if (!transitionToMini && !transitionToNormal) //############  NO TRANSITION
            {
                if (delta > 0) // Enlarging stream...
                {
                    if (_sectorRecycle)
                        freeList = FindFreeSectors(newSectorType); // Collect available free sectors

                    // ReSharper disable once UnusedVariable
                    sv = new StreamView(sectorChain, newSectorSize, length, freeList, SourceStream);

                    //Set up  destination chain
                    SetSectorChain(sectorChain);
                }
                else if (delta < 0) // Reducing size...
                {
                    var nSec = (int)Math.Floor(((double)(Math.Abs(delta)) /
                                                 newSectorSize)); //number of sectors to mark as free

                    if (newSectorSize == Sector.MINISECTOR_SIZE)
                        FreeMiniChain(sectorChain, nSec, _eraseFreeSectors);
                    else
                        FreeChain(sectorChain, nSec, _eraseFreeSectors);
                }

                if (sectorChain.Count > 0)
                {
                    cfItem.DirEntry.StartSetc = sectorChain[0].Id;
                    cfItem.DirEntry.Size = length;
                }
                else
                {
                    cfItem.DirEntry.StartSetc = Sector.ENDOFCHAIN;
                    cfItem.DirEntry.Size = 0;
                }
            }
            else if (transitionToMini) //############## TRANSITION TO MINISTREAM
            {
                // Transition Normal chain -> Mini chain

                // Collect available MINI free sectors

                if (_sectorRecycle)
                    freeList = FindFreeSectors(SectorType.Mini);

                sv = new StreamView(oldChain, oldSectorSize, oldSize, null, SourceStream);

                // Reset start sector and size of dir entry
                cfItem.DirEntry.StartSetc = Sector.ENDOFCHAIN;
                cfItem.DirEntry.Size = 0;

                var newChain = GetMiniSectorChain(Sector.ENDOFCHAIN);
                var destSv = new StreamView(newChain, Sector.MINISECTOR_SIZE, length, freeList, SourceStream);

                // Buffered trimmed copy from old (larger) to new (smaller)
                var cnt = 4096 < length ? 4096 : (int)length;

                var buf = new byte[4096];
                var toRead = length;

                //Copy old to new chain
                while (toRead > cnt)
                {
                    cnt = sv.Read(buf, 0, cnt);
                    toRead -= cnt;
                    destSv.Write(buf, 0, cnt);
                }

                sv.Read(buf, 0, (int)toRead);
                destSv.Write(buf, 0, (int)toRead);

                //Free old chain
                FreeChain(oldChain, _eraseFreeSectors);

                //Set up destination chain
                AllocateMiniSectorChain(destSv.BaseSectorChain);

                // Persist to normal strea
                PersistMiniStreamToStream(destSv.BaseSectorChain);

                //Update dir item
                if (destSv.BaseSectorChain.Count > 0)
                {
                    cfItem.DirEntry.StartSetc = destSv.BaseSectorChain[0].Id;
                    cfItem.DirEntry.Size = length;
                }
                else
                {
                    cfItem.DirEntry.StartSetc = Sector.ENDOFCHAIN;
                    cfItem.DirEntry.Size = 0;
                }
            }
            else //############## TRANSITION TO NORMAL STREAM
            {
                // Transition Mini chain -> Normal chain

                if (_sectorRecycle)
                    freeList = FindFreeSectors(SectorType.Normal); // Collect available Normal free sectors

                sv = new StreamView(oldChain, oldSectorSize, oldSize, null, SourceStream);

                var newChain = GetNormalSectorChain(Sector.ENDOFCHAIN);
                var destSv = new StreamView(newChain, GetSectorSize(), length, freeList, SourceStream);

                var cnt = 256 < length ? 256 : (int)length;

                var buf = new byte[256];
                var toRead = Math.Min(length, cfItem.Size);

                //Copy old to new chain
                while (toRead > cnt)
                {
                    cnt = sv.Read(buf, 0, cnt);
                    toRead -= cnt;
                    destSv.Write(buf, 0, cnt);
                }

                sv.Read(buf, 0, (int)toRead);
                destSv.Write(buf, 0, (int)toRead);

                //Free old mini chain
                // ReSharper disable once UnusedVariable
                var oldChainCount = oldChain.Count;
                FreeMiniChain(oldChain, _eraseFreeSectors);

                //Set up normal destination chain
                AllocateSectorChain(destSv.BaseSectorChain);

                //Update dir item
                if (destSv.BaseSectorChain.Count > 0)
                {
                    cfItem.DirEntry.StartSetc = destSv.BaseSectorChain[0].Id;
                    cfItem.DirEntry.Size = length;
                }
                else
                {
                    cfItem.DirEntry.StartSetc = Sector.ENDOFCHAIN;
                    cfItem.DirEntry.Size = 0;
                }
            }
        }

        internal void WriteData(CFItem cfItem, long position, byte[] buffer)
        {
            WriteData(cfItem, buffer, position, 0, buffer.Length);
        }

        internal void WriteData(CFItem cfItem, byte[] buffer, long position, int offset, int count)
        {
            if (buffer == null)
                throw new CFInvalidOperation("Parameter [buffer] cannot be null");

            if (cfItem.DirEntry == null)
                throw new CFException("Internal error [cfItem.DirEntry] cannot be null");

            if (buffer.Length == 0) return;

            // Get delta size induced by client
            var delta = (position + count) - cfItem.Size < 0 ? 0 : (position + count) - cfItem.Size;
            var newLength = cfItem.Size + delta;

            SetStreamLength(cfItem, newLength);

            // Calculate NEW sectors SIZE
            var st = SectorType.Normal;
            var sectorSize = GetSectorSize();

            if (cfItem.Size < _header.MinSizeStandardStream)
            {
                st = SectorType.Mini;
                sectorSize = Sector.MINISECTOR_SIZE;
            }

            var sectorChain = GetSectorChain(cfItem.DirEntry.StartSetc, st);
            var sv = new StreamView(sectorChain, sectorSize, newLength, null, SourceStream);

            sv.Seek(position, SeekOrigin.Begin);
            sv.Write(buffer, offset, count);

            if (cfItem.Size < _header.MinSizeStandardStream)
            {
                PersistMiniStreamToStream(sv.BaseSectorChain);
                //SetSectorChain(sv.BaseSectorChain);
            }
        }

        internal void WriteData(CFItem cfItem, byte[] buffer)
        {
            WriteData(cfItem, 0, buffer);
        }

        /// <summary>
        /// Check file size limit ( 2GB for version 3 )
        /// </summary>
        internal void CheckFileLength()
        {
            throw new NotImplementedException();
        }

        internal int ReadData(CFStream cFStream, long position, byte[] buffer, int count)
        {
            if (count > buffer.Length)
                throw new ArgumentException("count parameter exceeds buffer size");

            var de = cFStream.DirEntry;

            count = (int)Math.Min(de.Size - position, count);

            StreamView sView;

            if (de.Size < _header.MinSizeStandardStream)
            {
                sView
                    = new StreamView(GetSectorChain(de.StartSetc, SectorType.Mini), Sector.MINISECTOR_SIZE, de.Size,
                        null, SourceStream);
            }
            else
            {
                sView = new StreamView(GetSectorChain(de.StartSetc, SectorType.Normal), GetSectorSize(), de.Size, null,
                    SourceStream);
            }

            sView.Seek(position, SeekOrigin.Begin);
            var result = sView.Read(buffer, 0, count);

            return result;
        }

        internal int ReadData(CFStream cFStream, long position, byte[] buffer, int offset, int count)
        {
            var de = cFStream.DirEntry;

            count = (int)Math.Min(de.Size - offset, count);

            StreamView sView;

            if (de.Size < _header.MinSizeStandardStream)
            {
                sView
                    = new StreamView(GetSectorChain(de.StartSetc, SectorType.Mini), Sector.MINISECTOR_SIZE, de.Size,
                        null, SourceStream);
            }
            else
            {
                sView = new StreamView(GetSectorChain(de.StartSetc, SectorType.Normal), GetSectorSize(), de.Size, null,
                    SourceStream);
            }

            sView.Seek(position, SeekOrigin.Begin);
            var result = sView.Read(buffer, offset, count);

            return result;
        }

        internal byte[] GetData(CFStream cFStream)
        {
            if (IsClosed)
                throw new CFDisposedException("Compound File closed: cannot access data");

            byte[] result;

            var de = cFStream.DirEntry;

            //IDirectoryEntry root = directoryEntries[0];

            if (de.Size < _header.MinSizeStandardStream)
            {
                var miniView
                    = new StreamView(GetSectorChain(de.StartSetc, SectorType.Mini), Sector.MINISECTOR_SIZE, de.Size,
                        null, SourceStream);

                var br = new BinaryReader(miniView);

                result = br.ReadBytes((int)de.Size);
                br.Close();
            }
            else
            {
                var sView
                    = new StreamView(GetSectorChain(de.StartSetc, SectorType.Normal), GetSectorSize(), de.Size, null,
                        SourceStream);

                result = new byte[(int)de.Size];

                sView.Read(result, 0, result.Length);
            }

            return result;
        }

        public byte[] GetDataBySID(int sid)
        {
            if (IsClosed)
                throw new CFDisposedException("Compound File closed: cannot access data");
            if (sid < 0)
                return null;
            byte[] result;
            try
            {
                var de = _directoryEntries[sid];
                if (de.Size < _header.MinSizeStandardStream)
                {
                    var miniView
                        = new StreamView(GetSectorChain(de.StartSetc, SectorType.Mini), Sector.MINISECTOR_SIZE, de.Size,
                            null, SourceStream);
                    var br = new BinaryReader(miniView);
                    result = br.ReadBytes((int)de.Size);
                    br.Close();
                }
                else
                {
                    var sView
                        = new StreamView(GetSectorChain(de.StartSetc, SectorType.Normal), GetSectorSize(), de.Size,
                            null, SourceStream);
                    result = new byte[(int)de.Size];
                    sView.Read(result, 0, result.Length);
                }
            }
            catch
            {
                throw new CFException("Cannot get data for SID");
            }
            return result;
        }

        public Guid GetGuidBySID(int sid)
        {
            if (IsClosed)
                throw new CFDisposedException("Compound File closed: cannot access data");
            if (sid < 0)
                throw new CFException("Invalid SID");
            var de = _directoryEntries[sid];
            return de.StorageCLSID;
        }

        public Guid GetGuidForStream(int sid)
        {
            if (IsClosed)
                throw new CFDisposedException("Compound File closed: cannot access data");
            if (sid < 0)
                throw new CFException("Invalid SID");
            var g = new Guid("00000000000000000000000000000000");
            //find first storage containing a non-zero CLSID before SID in directory structure
            for (var i = sid - 1; i >= 0; i--)
            {
                if (_directoryEntries[i].StorageCLSID != g && _directoryEntries[i].StgType == StgType.StgStorage)
                {
                    return _directoryEntries[i].StorageCLSID;
                }
            }
            return g;
        }

        private static int Ceiling(double d)
        {
            return (int)Math.Ceiling(d);
        }

        private static int LowSaturation(int i)
        {
            return i > 0 ? i : 0;
        }

        internal void InvalidateDirectoryEntry(int sid)
        {
            if (sid >= _directoryEntries.Count)
                throw new CFException("Invalid SID of the directory entry to remove");

            //Random r = new Random();
            _directoryEntries[sid].SetEntryName("_DELETED_NAME_" + sid);
            _directoryEntries[sid].StgType = StgType.StgInvalid;
        }

        internal void FreeAssociatedData(int sid)
        {
            // Clear the associated stream (or ministream) if required
            if (_directoryEntries[sid].Size > 0) //thanks to Mark Bosold for this !
            {
                if (_directoryEntries[sid].Size < _header.MinSizeStandardStream)
                {
                    var miniChain
                        = GetSectorChain(_directoryEntries[sid].StartSetc, SectorType.Mini);
                    FreeMiniChain(miniChain, _eraseFreeSectors);
                }
                else
                {
                    var chain
                        = GetSectorChain(_directoryEntries[sid].StartSetc, SectorType.Normal);
                    FreeChain(chain, _eraseFreeSectors);
                }
            }
        }

        /// <summary>
        /// Close the Compound File object <see cref="T:OpenMcdf.CompoundFile">CompoundFile</see> and
        /// free all associated resources (e.g. open file handle and allocated memory).
        /// <remarks>
        /// When the <see cref="T:OpenMcdf.CompoundFile.Close()">Close</see> method is called,
        /// all the associated stream and storage objects are invalidated:
        /// any operation invoked on them will produce a <see cref="T:OpenMcdf.CFDisposedException">CFDisposedException</see>.
        /// </remarks>
        /// </summary>
        /// <example>
        /// <code>
        ///    const string FILENAME = "CompoundFile.cfs";
        ///    CompoundFile cf = new CompoundFile(FILENAME);
        ///
        ///    CFStorage st = cf.RootStorage.GetStorage("MyStorage");
        ///    cf.Close();
        ///
        ///    try
        ///    {
        ///        byte[] temp = st.GetStream("MyStream").GetData();
        ///
        ///        // The following line will fail because back-end object has been closed
        ///        Assert.Fail("Stream without media");
        ///    }
        ///    catch (Exception ex)
        ///    {
        ///        Assert.IsTrue(ex is CFDisposedException);
        ///    }
        /// </code>
        /// </example>
        public void Close()
        {
#pragma warning disable CS0618 // Type or member is obsolete
            Close(true);
#pragma warning restore CS0618 // Type or member is obsolete
        }

        private bool _closeStream = true;

        [Obsolete("Use flag LeaveOpen in CompoundFile constructor")]
        public void Close(bool closeStream)
        {
            _closeStream = closeStream;
            ((IDisposable)this).Dispose();
        }

        #region IDisposable Members

        void IDisposable.Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion IDisposable Members

        /// <summary>
        /// When called from user code, release all resources, otherwise, in the case runtime called it,
        /// only unmanagd resources are released.
        /// </summary>
        /// <param name="disposing">If true, method has been called from User code, if false it's been called from .net runtime</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!IsClosed)
            {
                lock (_lockObject)
                {
                    if (disposing)
                    {
                        // Call from user code...

                        if (_sectors != null)
                        {
                            _sectors.Clear();
                            _sectors = null;
                        }

                        RootStorage = null; // Some problem releasing resources...
                        _header = null;
                        _directoryEntries.Clear();
                        _directoryEntries = null;
                        FileName = null;
                        //this.lockObject = null;
#if !FLAT_WRITE
                            this.buffer = null;
#endif
                    }

                    if (SourceStream != null && _closeStream &&
                        !Configuration.HasFlag(CFSConfiguration.LeaveOpen))
                        SourceStream.Close();
                }
            }

            IsClosed = true;
        }

        internal IList<IDirectoryEntry> GetDirectories()
        {
            return _directoryEntries;
        }

        private IEnumerable<IDirectoryEntry> FindDirectoryEntries(string entryName)
        {
            return _directoryEntries.Where(d => d.GetEntryName() == entryName && d.StgType != StgType.StgInvalid).ToList();
        }

        /// <summary>
        /// Get a list of all entries with a given name contained in the document.
        /// </summary>
        /// <param name="entryName">Name of entries to retrieve</param>
        /// <returns>A list of name-matching entries</returns>
        /// <remarks>This function is aimed to speed up entity lookup in
        /// flat-structure files (only one or little more known entries)
        /// without the performance penalty related to entities hierarchy constraints.
        /// There is no implied hierarchy in the returned list.
        /// </remarks>
        public IList<CFItem> GetAllNamedEntries(string entryName)
        {
            var dirEntries = FindDirectoryEntries(entryName);

            return (from id in dirEntries
                    where id.GetEntryName() == entryName && id.StgType != StgType.StgInvalid
                    select id.StgType == StgType.StgStorage
                        ? new CFStorage(this, id)
                        : (CFItem)new CFStream(this, id)).ToList();
        }

        public int GetNumDirectories()
        {
            if (IsClosed)
                throw new CFDisposedException("Compound File closed: cannot access data");
            return _directoryEntries.Count;
        }

        public string GetNameDirEntry(int id)
        {
            if (IsClosed)
                throw new CFDisposedException("Compound File closed: cannot access data");
            if (id < 0)
                throw new CFException("Invalid Storage ID");
            return _directoryEntries[id].Name;
        }

        public StgType GetStorageType(int id)
        {
            if (IsClosed)
                throw new CFDisposedException("Compound File closed: cannot access data");
            if (id < 0)
                throw new CFException("Invalid Storage ID");
            return _directoryEntries[id].StgType;
        }

        /// <summary>
        /// Compress free space by removing unallocated sectors from compound file
        /// effectively reducing stream or file size.
        /// </summary>
        /// <remarks>
        /// Current implementation supports compression only for ver. 3 compound files.
        /// </remarks>
        /// <example>
        /// <code>
        ///
        ///  //This code has been extracted from unit test
        ///
        ///    string FILENAME = "MultipleStorage3.cfs";
        ///
        ///    FileInfo srcFile = new FileInfo(FILENAME);
        ///
        ///    File.Copy(FILENAME, "MultipleStorage_Deleted_Compress.cfs", true);
        ///
        ///    CompoundFile cf = new CompoundFile("MultipleStorage_Deleted_Compress.cfs", UpdateMode.Update, true, true);
        ///
        ///    CFStorage st = cf.RootStorage.GetStorage("MyStorage");
        ///    st = st.GetStorage("AnotherStorage");
        ///
        ///    Assert.IsNotNull(st);
        ///    st.Delete("Another2Stream"); //17Kb
        ///    cf.Commit();
        ///    cf.Close();
        ///
        ///    CompoundFile.ShrinkCompoundFile("MultipleStorage_Deleted_Compress.cfs");
        ///
        ///    FileInfo dstFile = new FileInfo("MultipleStorage_Deleted_Compress.cfs");
        ///
        ///    Assert.IsTrue(srcFile.Length > dstFile.Length);
        ///
        /// </code>
        /// </example>
        public static void ShrinkCompoundFile(Stream stream)
        {
            var cf = new CompoundFile(stream, CFSUpdateMode.Update, CFSConfiguration.LeaveOpen);

            if (cf._header.MajorVersion != (ushort)CFSVersion.Ver_3)
                throw new CFException(
                    "Current implementation of free space compression does not support version 4 of Compound File Format");

            using (var tempCF = new CompoundFile((CFSVersion)cf._header.MajorVersion, cf.Configuration))
            {
                //Copy Root CLSID
                tempCF.RootStorage.CLSID = new Guid(cf.RootStorage.CLSID.ToByteArray());

                DoCompression(cf.RootStorage, tempCF.RootStorage);

                //This could be a problem for v4
                using (var tmpMs = new MemoryStream((int)cf.SourceStream.Length))
                {
                    tempCF.Save(tmpMs);
                    tempCF.Close();

                    // If we were based on a writable stream, we update
                    // the stream and do reload from the compressed one...

                    stream.Seek(0, SeekOrigin.Begin);
                    tmpMs.WriteTo(stream);

                    stream.Seek(0, SeekOrigin.Begin);
                    stream.SetLength(tmpMs.Length);

                    tmpMs.Close();
                }

                cf.Close();
            }
        }

        /// <summary>
        /// Remove unallocated sectors from compound file in order to reduce its size.
        /// </summary>
        /// <remarks>
        /// Current implementation supports compression only for ver. 3 compound files.
        /// </remarks>
        /// <example>
        /// <code>
        ///
        ///  //This code has been extracted from unit test
        ///
        ///    string FILENAME = "MultipleStorage3.cfs";
        ///
        ///    FileInfo srcFile = new FileInfo(FILENAME);
        ///
        ///    File.Copy(FILENAME, "MultipleStorage_Deleted_Compress.cfs", true);
        ///
        ///    CompoundFile cf = new CompoundFile("MultipleStorage_Deleted_Compress.cfs", UpdateMode.Update, true, true);
        ///
        ///    CFStorage st = cf.RootStorage.GetStorage("MyStorage");
        ///    st = st.GetStorage("AnotherStorage");
        ///
        ///    Assert.IsNotNull(st);
        ///    st.Delete("Another2Stream"); //17Kb
        ///    cf.Commit();
        ///    cf.Close();
        ///
        ///    CompoundFile.ShrinkCompoundFile("MultipleStorage_Deleted_Compress.cfs");
        ///
        ///    FileInfo dstFile = new FileInfo("MultipleStorage_Deleted_Compress.cfs");
        ///
        ///    Assert.IsTrue(srcFile.Length > dstFile.Length);
        ///
        /// </code>
        /// </example>
        public static void ShrinkCompoundFile(string fileName)
        {
            using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite))
            {
                ShrinkCompoundFile(fs);
                fs.Close();
            }
        }

        /// <summary>
        /// Recursively clones valid structures, avoiding to copy free sectors.
        /// </summary>
        /// <param name="currSrcStorage">Current source storage to clone</param>
        /// <param name="currDstStorage">Current cloned destination storage</param>
        private static void DoCompression(CFStorage currSrcStorage, CFStorage currDstStorage)
        {
            void Va(CFItem item)
            {
                if (item.IsStream)
                {
                    var itemAsStream = item as CFStream;
                    if (itemAsStream == null)
                        return;

                    var st = currDstStorage.AddStream(itemAsStream.Name);
                    st.SetData(itemAsStream.GetData());
                }
                else if (item.IsStorage)
                {
                    var itemAsStorage = item as CFStorage;
                    if (itemAsStorage == null)
                        return;

                    var strg = currDstStorage.AddStorage(itemAsStorage.Name);
                    strg.CLSID = new Guid(itemAsStorage.CLSID.ToByteArray());
                    DoCompression(itemAsStorage, strg); // recursion, one level deeper
                }
            }

            currSrcStorage.VisitEntries(Va, false);
        }
    }
}