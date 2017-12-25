using System;

namespace OpenMcdf.PerfTest
{
    public static class Helpers
    {
        public static byte[] GetBuffer(int count)
        {
            Random r = new Random();
            byte[] b = new byte[count];
            r.NextBytes(b);
            return b;
        }

        public static byte[] GetBuffer(int count, byte c)
        {
            byte[] b = new byte[count];
            for (int i = 0; i < b.Length; i++)
            {
                b[i] = c;
            }

            return b;
        }

        public static bool CompareBuffer(byte[] b, byte[] p)
        {
            if (b == null && p == null)
                throw new Exception("Null buffers");

            if (b == null && p != null) 
                return false;

            if (b != null && p == null) 
                return false;

            if (b.Length != p.Length)
                return false;

            for (int i = 0; i < b.Length; i++)
            {
                if (b[i] != p[i])
                    return false;
            }

            return true;
        }

        internal static void CreateFile(string fileName)
        {
            const int MAX_STREAM_COUNT = 5000;

            CompoundFile cf = new CompoundFile();
            for (int i = 0; i < MAX_STREAM_COUNT; i++)
            {
                cf.RootStorage.AddStream("Test" + i.ToString()).SetData(Helpers.GetBuffer(300));
            }
            cf.Save(fileName);
            cf.Close();
        }
    }
}
