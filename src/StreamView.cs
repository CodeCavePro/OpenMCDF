/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/


using System;
using System.Collections.Generic;
using System.IO;

namespace OpenMcdf
{
    /// <inheritdoc />
    /// <summary>
    /// Stream decorator for a Sector or miniSector chain
    /// </summary>
    internal class StreamView : Stream
    {
        private readonly int _sectorSize;

        private long _position;

        private readonly Stream _stream;

        private readonly bool _isFatStream;

        public StreamView(IList<Sector> sectorChain, int sectorSize, Stream stream)
        {
            if (sectorSize <= 0)
                throw new CFException("Sector size must be greater than zero");

            BaseSectorChain = sectorChain ?? throw new CFException("Sector Chain cannot be null");
            _sectorSize = sectorSize;
            _stream = stream;
        }

        public StreamView(IList<Sector> sectorChain, int sectorSize, long length, Queue<Sector> availableSectors,
            Stream stream, bool isFatStream = false)
            : this(sectorChain, sectorSize, stream)
        {
            _isFatStream = isFatStream;
            AdjustLength(length, availableSectors);
        }

        public IEnumerable<Sector> FreeSectors { get; } = new List<Sector>();

        public IList<Sector> BaseSectorChain { get; }

        public override bool CanRead => true;

        public override bool CanSeek => true;

        public override bool CanWrite => true;

        public override void Flush()
        {
        }

        private long _length;

        public override long Length => _length;

        public override long Position
        {
            get => _position;

            set
            {
                if (_position > _length - 1)
                    throw new ArgumentOutOfRangeException(nameof(value));

                _position = value;
            }
        }

#if !NETSTANDARD1_6 && !NETSTANDARD2_0
        public override void Close()
        {
            base.Close();
        }
#endif

        private byte[] _buf = new byte[4];

        public int ReadInt32()
        {
            Read(_buf, 0, 4);
            return (((_buf[0] | (_buf[1] << 8)) | (_buf[2] << 16)) | (_buf[3] << 24));
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            var nRead = 0;
            int nToRead;

            if (BaseSectorChain == null || BaseSectorChain.Count <= 0)
                return 0;

            // First sector
            var secIndex = (int) (_position / _sectorSize);

            // Bytes to read count is the min between request count
            // and sector border

            nToRead = Math.Min(
                BaseSectorChain[0].Size - ((int) _position % _sectorSize),
                count);

            if (secIndex < BaseSectorChain.Count)
            {
                Buffer.BlockCopy(
                    BaseSectorChain[secIndex].GetData(),
                    (int) (_position % _sectorSize),
                    buffer,
                    offset,
                    nToRead
                );
            }

            nRead += nToRead;

            secIndex++;

            // Central sectors
            while (nRead < (count - _sectorSize))
            {
                nToRead = _sectorSize;

                Buffer.BlockCopy(
                    BaseSectorChain[secIndex].GetData(),
                    0,
                    buffer,
                    offset + nRead,
                    nToRead
                );

                nRead += nToRead;
                secIndex++;
            }

            // Last sector
            nToRead = count - nRead;

            if (nToRead != 0)
            {
                Buffer.BlockCopy(
                    BaseSectorChain[secIndex].GetData(),
                    0,
                    buffer,
                    offset + nRead,
                    nToRead
                );

                nRead += nToRead;
            }

            _position += nRead;

            return nRead;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            switch (origin)
            {
                case SeekOrigin.Begin:
                    _position = offset;
                    break;

                case SeekOrigin.Current:
                    _position += offset;
                    break;

                case SeekOrigin.End:
                    _position = Length - offset;
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(origin), origin, null);
            }

            AdjustLength(_position);

            return _position;
        }

        private void AdjustLength(long value, Queue<Sector> availableSectors = null)
        {
            _length = value;

            var delta = value - (BaseSectorChain.Count * (long) _sectorSize);

            if (delta <= 0)
                return;

            // enlargement required
            var nSec = (int) Math.Ceiling(((double) delta / _sectorSize));

            while (nSec > 0)
            {
                Sector t;

                if (availableSectors == null || availableSectors.Count == 0)
                {
                    t = new Sector(_sectorSize, _stream);

                    if (_sectorSize == Sector.MINISECTOR_SIZE)
                        t.Type = SectorType.Mini;
                }
                else
                {
                    t = availableSectors.Dequeue();
                }

                if (_isFatStream)
                {
                    t.InitFATData();
                }
                BaseSectorChain.Add(t);
                nSec--;
            }
        }

        public override void SetLength(long value)
        {
            AdjustLength(value);
        }

        public void WriteInt32(int val)
        {
            var buffer = new byte[4];
            buffer[0] = (byte) val;
            buffer[1] = (byte) (val << 8);
            buffer[2] = (byte) (val << 16);
            buffer[3] = (byte) (val << 32);
            Write(buffer, 0, 4);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            var byteWritten = 0;
            int roundByteWritten;

            // Assure length
            if ((_position + count) > _length)
                AdjustLength((_position + count));

            if (BaseSectorChain == null)
                return;

            // First sector
            var secOffset = (int) (_position / _sectorSize);
            var secShift = (int) _position % _sectorSize;

            roundByteWritten = Math.Min(_sectorSize - (int) (_position % _sectorSize), count);

            if (secOffset < BaseSectorChain.Count)
            {
                Buffer.BlockCopy(
                    buffer,
                    offset,
                    BaseSectorChain[secOffset].GetData(),
                    secShift,
                    roundByteWritten
                );

                BaseSectorChain[secOffset].DirtyFlag = true;
            }

            byteWritten += roundByteWritten;
            offset += roundByteWritten;
            secOffset++;

            // Central sectors
            while (byteWritten < (count - _sectorSize))
            {
                roundByteWritten = _sectorSize;

                Buffer.BlockCopy(
                    buffer,
                    offset,
                    BaseSectorChain[secOffset].GetData(),
                    0,
                    roundByteWritten
                );

                BaseSectorChain[secOffset].DirtyFlag = true;

                byteWritten += roundByteWritten;
                offset += roundByteWritten;
                secOffset++;
            }

            // Last sector
            roundByteWritten = count - byteWritten;

            if (roundByteWritten != 0)
            {
                Buffer.BlockCopy(
                    buffer,
                    offset,
                    BaseSectorChain[secOffset].GetData(),
                    0,
                    roundByteWritten
                );

                BaseSectorChain[secOffset].DirtyFlag = true;

                offset += roundByteWritten;
                byteWritten += roundByteWritten;
            }

            _position += count;
        }
    }
}