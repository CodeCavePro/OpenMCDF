using System;
using System.IO;
using NUnit.Framework;

namespace OpenMcdf.Test
{
    [TestFixture]
    public class StreamRWTest
    {
        [Test]
        public void ReadInt64_MaxSizeRead()
        {
            Int64 input = Int64.MaxValue;
            byte[] bytes = BitConverter.GetBytes(input);
            long actual = 0;
            using (MemoryStream memStream = new MemoryStream(bytes))
            {
                OpenMcdf.StreamRW reader = new OpenMcdf.StreamRW(memStream);
                actual = reader.ReadInt64();
            }
            Assert.AreEqual((long) input, actual);
        }

        [Test]
        public void ReadInt64_SmallNumber()
        {
            Int64 input = 1234;
            byte[] bytes = BitConverter.GetBytes(input);
            long actual = 0;
            using (MemoryStream memStream = new MemoryStream(bytes))
            {
                OpenMcdf.StreamRW reader = new OpenMcdf.StreamRW(memStream);
                actual = reader.ReadInt64();
            }
            Assert.AreEqual((long) input, actual);
        }

        [Test]
        public void ReadInt64_Int32MaxPlusTen()
        {
            Int64 input = (Int64) Int32.MaxValue + 10;
            byte[] bytes = BitConverter.GetBytes(input);
            long actual = 0;
            using (MemoryStream memStream = new MemoryStream(bytes))
            {
                OpenMcdf.StreamRW reader = new OpenMcdf.StreamRW(memStream);
                actual = reader.ReadInt64();
            }
            Assert.AreEqual((long) input, actual);
        }
    }
}