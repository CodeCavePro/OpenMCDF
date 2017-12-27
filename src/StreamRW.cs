/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/

using System.IO;

namespace OpenMcdf
{
    internal class StreamRW
    {
        private readonly byte[] _buffer;
        private readonly Stream _stream;

        public StreamRW(Stream stream)
        {
            _stream = stream;
            _buffer = new byte[8];
        }

        public long Seek(long offset)
        {
            return _stream.Seek(offset, SeekOrigin.Begin);
        }

        public byte ReadByte()
        {
            _stream.Read(_buffer, 0, 1);
            return _buffer[0];
        }

        public ushort ReadUInt16()
        {
            _stream.Read(_buffer, 0, 2);
            return (ushort) (_buffer[0] | (_buffer[1] << 8));
        }

        public int ReadInt32()
        {
            _stream.Read(_buffer, 0, 4);
            return _buffer[0] | (_buffer[1] << 8) | (_buffer[2] << 16) | (_buffer[3] << 24);
        }

        public uint ReadUInt32()
        {
            _stream.Read(_buffer, 0, 4);
            return (uint) (_buffer[0] | (_buffer[1] << 8) | (_buffer[2] << 16) | (_buffer[3] << 24));
        }

        public long ReadInt64()
        {
            _stream.Read(_buffer, 0, 8);
            var ls = (uint) (_buffer[0] | (_buffer[1] << 8) | (_buffer[2] << 16) | (_buffer[3] << 24));
            var ms = (uint) ((_buffer[4]) | (_buffer[5] << 8) | (_buffer[6] << 16) | (_buffer[7] << 24));
            return (long) (((ulong) ms << 32) | ls);
        }

        public ulong ReadUInt64()
        {
            _stream.Read(_buffer, 0, 8);
            return (ulong) (_buffer[0] | (_buffer[1] << 8) | (_buffer[2] << 16) | (_buffer[3] << 24) | (_buffer[4] << 32) |
                            (_buffer[5] << 40) | (_buffer[6] << 48) | (_buffer[7] << 56));
        }

        public byte[] ReadBytes(int count)
        {
            var result = new byte[count];
            _stream.Read(result, 0, count);
            return result;
        }

        public byte[] ReadBytes(int count, out int rCount)
        {
            var result = new byte[count];
            rCount = _stream.Read(result, 0, count);
            return result;
        }

        public void Write(byte b)
        {
            _buffer[0] = b;
            _stream.Write(_buffer, 0, 1);
        }

        public void Write(ushort value)
        {
            _buffer[0] = (byte) value;
            _buffer[1] = (byte) (value >> 8);

            _stream.Write(_buffer, 0, 2);
        }

        public void Write(int value)
        {
            _buffer[0] = (byte) value;
            _buffer[1] = (byte) (value >> 8);
            _buffer[2] = (byte) (value >> 16);
            _buffer[3] = (byte) (value >> 24);

            _stream.Write(_buffer, 0, 4);
        }

        public void Write(long value)
        {
            _buffer[0] = (byte) value;
            _buffer[1] = (byte) (value >> 8);
            _buffer[2] = (byte) (value >> 16);
            _buffer[3] = (byte) (value >> 24);
            _buffer[4] = (byte) (value >> 32);
            _buffer[5] = (byte) (value >> 40);
            _buffer[6] = (byte) (value >> 48);
            _buffer[7] = (byte) (value >> 56);

            _stream.Write(_buffer, 0, 8);
        }

        public void Write(uint value)
        {
            _buffer[0] = (byte) value;
            _buffer[1] = (byte) (value >> 8);
            _buffer[2] = (byte) (value >> 16);
            _buffer[3] = (byte) (value >> 24);

            _stream.Write(_buffer, 0, 4);
        }

        public void Write(byte[] value)
        {
            _stream.Write(value, 0, value.Length);
        }

        public void Close()
        {
            //Nothing to do ;-)
        }
    }
}