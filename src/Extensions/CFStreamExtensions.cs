using System;
using System.IO;

namespace OpenMcdf.Extensions
{
    public static class CFStreamExtension
    {
        private class StreamDecorator : Stream
        {
            private CFStream _cfStream;
            private long _position;

            public StreamDecorator(CFStream cfstream)
            {
                _cfStream = cfstream;
            }

            public override bool CanRead => true;

            public override bool CanSeek => true;

            public override bool CanWrite => true;

            public override void Flush()
            {
                // nothing to do;
            }

            public override long Length => _cfStream.Size;

            public override long Position
            {
                get => _position;
                set => _position = value;
            }

            public override int Read(byte[] buffer, int offset, int count)
            {
                if (count > buffer.Length)
                    throw new ArgumentException("Count parameter exceeds buffer size");

                if (buffer == null)
                    throw new ArgumentNullException("Buffer cannot be null");

                if (offset < 0 || count < 0)
                    throw new ArgumentOutOfRangeException("Offset and Count parameters must be non-negative numbers");

                if (_position >= _cfStream.Size)
                    return 0;

                count = _cfStream.Read(buffer, _position, offset, count);
                _position += count;
                return count;
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
                        _position -= offset;
                        break;
                    default:
                        throw new Exception("Invalid origin selected");
                }

                return _position;
            }

            public override void SetLength(long value)
            {
                _cfStream.Resize(value);
            }

            public override void Write(byte[] buffer, int offset, int count)
            {
                _cfStream.Write(buffer, _position, offset, count);
                _position += count;
            }

#if !NETSTANDARD1_6 && !NETSTANDARD2_0
            public override void Close()
            {
                // Do nothing
            }
#endif
        }

        /// <summary>
        /// Return the current <see cref="T:OpenMcdf.CFStream">CFStream</see> object 
        /// as a <see cref="T:System.IO.Stream">Stream</see> object.
        /// </summary>
        /// <param name="cfStream">Current <see cref="T:OpenMcdf.CFStream">CFStream</see> object</param>
        /// <returns>A <see cref="T:System.IO.Stream">Stream</see> object representing structured stream data</returns>
        public static Stream AsIoStream(this CFStream cfStream)
        {
            return new StreamDecorator(cfStream);
        }

        /// <summary>
        /// Return the current <see cref="T:OpenMcdf.CFStream">CFStream</see> object 
        /// as a OLE properties Stream.
        /// </summary>
        /// <param name="cfStream"></param>
        /// <returns>A <see cref="T:OpenMcdf.Extensions.OLEProperties.PropertySetStream">OLE Propertie stream</see></returns>
        public static OLEProperties.PropertySetStream AsOleProperties(this CFStream cfStream)
        {
            var result = new OLEProperties.PropertySetStream();
            result.Read(new BinaryReader(new StreamDecorator(cfStream)));
            return result;
        }
    }
}