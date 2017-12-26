using System;
using System.IO;
using NUnit.Framework;
using OpenMcdf.Extensions;

namespace OpenMcdf.Test
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestFixture]
    public class CFSStreamExtensionsTest
    {
        [Test]
        public void Test_AS_IOSTREAM_READ()
        {
            CompoundFile cf = new CompoundFile("MultipleStorage.cfs");

            Stream s = cf.RootStorage.GetStorage("MyStorage").GetStream("MyStream").AsIOStream();
            BinaryReader br = new BinaryReader(s);
            byte[] result = br.ReadBytes(32);
            Assert.IsTrue(Helpers.CompareBuffer(Helpers.GetBuffer(32, 1), result));
        }

        [Test]
        public void Test_AS_IOSTREAM_WRITE()
        {
            const String cmp = "Hello World of BinaryWriter !";

            CompoundFile cf = new CompoundFile();
            Stream s = cf.RootStorage.AddStream("ANewStream").AsIOStream();
            BinaryWriter bw = new BinaryWriter(s);
            bw.Write(cmp);
            cf.Save("$ACFFile.cfs");
            cf.Close();

            cf = new CompoundFile("$ACFFile.cfs");
            BinaryReader br = new BinaryReader(cf.RootStorage.GetStream("ANewStream").AsIOStream());
            String st = br.ReadString();
            Assert.IsTrue(st == cmp);
            cf.Close();
        }
    }
}
