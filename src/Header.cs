/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/


using System.IO;
using System.Linq;

namespace OpenMcdf
{
    internal class Header
    {
        //0 8 Compound document file identifier: D0H CFH 11H E0H A1H B1H 1AH E1H

        public byte[] HeaderSignature { get; private set; }

        //8 16 Unique identifier (UID) of this file (not of interest in the following, may be all 0)

        public byte[] CLSID { get; set; }

        //24 2 Revision number of the file format (most used is 003EH)

        public ushort MinorVersion { get; private set; }

        //26 2 Version number of the file format (most used is 0003H)

        public ushort MajorVersion { get; private set; }

        //28 2 Byte order identifier (➜4.2): FEH FFH = Little-Endian FFH FEH = Big-Endian

        public ushort ByteOrder { get; private set; }

        //30 2 Size of a sector in the compound document file (➜3.1) in power-of-two (ssz), real sector
        //size is sec_size = 2ssz bytes (minimum value is 7 which means 128 bytes, most used 
        //value is 9 which means 512 bytes)

        public ushort SectorShift { get; private set; }

        //32 2 Size of a short-sector in the short-stream container stream (➜6.1) in power-of-two (sssz),
        //real short-sector size is short_sec_size = 2sssz bytes (maximum value is sector size
        //ssz, see above, most used value is 6 which means 64 bytes)

        public ushort MiniSectorShift { get; private set; }

        //34 10 Not used

        public byte[] UnUsed { get; private set; }

        //44 4 Total number of sectors used Directory (➜5.2)

        public int DirectorySectorsNumber { get; set; }

        //44 4 Total number of sectors used for the sector allocation table (➜5.2)

        public int FATSectorsNumber { get; set; }

        //48 4 SecID of first sector of the directory stream (➜7)

        public int FirstDirectorySectorId { get; set; }

        //52 4 Not used

        public uint UnUsed2 { get; private set; }

        //56 4 Minimum size of a standard stream (in bytes, minimum allowed and most used size is 4096
        //bytes), streams with an actual size smaller than (and not equal to) this value are stored as
        //short-streams (➜6)

        public uint MinSizeStandardStream { get; set; }

        //60 4 SecID of first sector of the short-sector allocation table (➜6.2), or –2 (End Of Chain
        //SecID, ➜3.1) if not extant

        /// <summary>
        /// This integer field contains the starting sector number for the mini FAT
        /// </summary>
        public int FirstMiniFATSectorId { get; set; }

        //64 4 Total number of sectors used for the short-sector allocation table (➜6.2)

        public uint MiniFATSectorsNumber { get; set; }

        //68 4 SecID of first sector of the master sector allocation table (➜5.1), or –2 (End Of Chain
        //SecID, ➜3.1) if no additional sectors used

        public int FirstDIFATSectorId { get; set; }

        //72 4 Total number of sectors used for the master sector allocation table (➜5.1)

        public uint DIFATSectorsNumber { get; set; }

        //76 436 First part of the master sector allocation table (➜5.1) containing 109 SecIDs

        public int[] DIFAT { get; }

        /// <summary>
        /// Structured Storage signature
        /// </summary>
        protected byte[] OleCFSSignature { get; }

        public Header()
            : this(3)
        {
        }

        public Header(ushort version)
        {
            DIFAT = new int[109];
            FirstDIFATSectorId = Sector.ENDOFCHAIN;
            FirstMiniFATSectorId = unchecked((int) 0xFFFFFFFE);
            MinSizeStandardStream = 4096;
            FirstDirectorySectorId = Sector.ENDOFCHAIN;
            UnUsed = new byte[6];
            MiniSectorShift = 6;
            SectorShift = 9;
            ByteOrder = 0xFFFE;
            MajorVersion = 0x0003;
            MinorVersion = 0x003E;
            CLSID = new byte[16];
            HeaderSignature = new byte[] {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};

            OleCFSSignature = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

            switch (version)
            {
                case 3:
                    MajorVersion = 3;
                    SectorShift = 0x0009;
                    break;

                case 4:
                    MajorVersion = 4;
                    SectorShift = 0x000C;
                    break;

                default:
                    throw new CFException("Invalid Compound File Format version");
            }

            for (var i = 0; i < 109; i++)
            {
                DIFAT[i] = Sector.FREESECT;
            }
        }

        public void Write(Stream stream)
        {
            var rw = new StreamRW(stream);

            rw.Write(HeaderSignature);
            rw.Write(CLSID);
            rw.Write(MinorVersion);
            rw.Write(MajorVersion);
            rw.Write(ByteOrder);
            rw.Write(SectorShift);
            rw.Write(MiniSectorShift);
            rw.Write(UnUsed);
            rw.Write(DirectorySectorsNumber);
            rw.Write(FATSectorsNumber);
            rw.Write(FirstDirectorySectorId);
            rw.Write(UnUsed2);
            rw.Write(MinSizeStandardStream);
            rw.Write(FirstMiniFATSectorId);
            rw.Write(MiniFATSectorsNumber);
            rw.Write(FirstDIFATSectorId);
            rw.Write(DIFATSectorsNumber);

            foreach (var i in DIFAT)
            {
                rw.Write(i);
            }

            if (MajorVersion == 4)
            {
                var zeroHead = new byte[3584];
                rw.Write(zeroHead);
            }

            rw.Close();
        }

        public void Read(Stream stream)
        {
            var rw = new StreamRW(stream);

            HeaderSignature = rw.ReadBytes(8);
            CheckSignature();
            CLSID = rw.ReadBytes(16);
            MinorVersion = rw.ReadUInt16();
            MajorVersion = rw.ReadUInt16();
            CheckVersion();
            ByteOrder = rw.ReadUInt16();
            SectorShift = rw.ReadUInt16();
            MiniSectorShift = rw.ReadUInt16();
            UnUsed = rw.ReadBytes(6);
            DirectorySectorsNumber = rw.ReadInt32();
            FATSectorsNumber = rw.ReadInt32();
            FirstDirectorySectorId = rw.ReadInt32();
            UnUsed2 = rw.ReadUInt32();
            MinSizeStandardStream = rw.ReadUInt32();
            FirstMiniFATSectorId = rw.ReadInt32();
            MiniFATSectorsNumber = rw.ReadUInt32();
            FirstDIFATSectorId = rw.ReadInt32();
            DIFATSectorsNumber = rw.ReadUInt32();

            for (var i = 0; i < 109; i++)
            {
                DIFAT[i] = rw.ReadInt32();
            }

            rw.Close();
        }

        private void CheckVersion()
        {
            if (MajorVersion != 3 && MajorVersion != 4)
                throw new CFFileFormatException(
                    "Unsupported Binary File Format version: OpenMcdf only supports Compound Files with major version equal to 3 or 4 ");
        }

        private void CheckSignature()
        {
            if (HeaderSignature.Where((t, i) => t != OleCFSSignature[i]).Any())
            {
                throw new CFFileFormatException("Invalid OLE structured storage file");
            }
        }
    }
}