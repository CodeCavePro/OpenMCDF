using System;
using System.Diagnostics;
using System.IO;
using HelpersFromTests = OpenMcdf.Test.Helpers;

namespace OpenMcdf.PerfTest
{
    public static class Helpers
    {
        internal static void StressMemory()
        {
            const int N_LOOP = 20;
            const int MB_SIZE = 10;

            byte[] b = HelpersFromTests.GetBuffer(1024 * 1024 * MB_SIZE); //2GB buffer
            byte[] cmp = new byte[] { 0x0, 0x1, 0x2, 0x3, 0x4, 0x5, 0x6, 0x7 };

            CompoundFile cf = new CompoundFile(CFSVersion.Ver_4, CFSConfiguration.Default);
            CFStream st = cf.RootStorage.AddStream("MySuperLargeStream");
            cf.Save("LARGE.cfs");
            cf.Close();

            //Console.WriteLine("Closed save");
            //Console.ReadKey();

            cf = new CompoundFile("LARGE.cfs", CFSUpdateMode.Update, CFSConfiguration.Default);
            CFStream cfst = cf.RootStorage.GetStream("MySuperLargeStream");

            Stopwatch sw = new Stopwatch();
            sw.Start();
            for (int i = 0; i < N_LOOP; i++)
            {

                cfst.Append(b);
                cf.Commit(true);

                Console.WriteLine("     Updated " + i.ToString());
                //Console.ReadKey();
            }

            cfst.Append(cmp);
            cf.Commit(true);
            sw.Stop();


            cf.Close();

            Console.WriteLine(sw.Elapsed.TotalMilliseconds);
            sw.Reset();

            //Console.WriteLine(sw.Elapsed.TotalMilliseconds);

            //Console.WriteLine("Closed Transacted");
            //Console.ReadKey();

            cf = new CompoundFile("LARGE.cfs");
            int count = 8;
            sw.Reset();
            sw.Start();
            byte[] data = new byte[count];
            count = cf.RootStorage.GetStream("MySuperLargeStream").Read(data, b.Length * (long)N_LOOP, count);
            sw.Stop();
            Console.Write(count);
            cf.Close();

            Console.WriteLine("Closed Final " + sw.ElapsedMilliseconds);
            Console.ReadKey();

        }

        internal static void DummyFile()
        {
            Console.WriteLine("Start");
            FileStream fs = new FileStream("myDummyFile", FileMode.Create);
            fs.Close();

            Stopwatch sw = new Stopwatch();

            byte[] b = HelpersFromTests.GetBuffer(1024 * 1024 * 50); //2GB buffer

            fs = new FileStream("myDummyFile", FileMode.Open);
            sw.Start();
            for (int i = 0; i < 42; i++)
            {

                fs.Seek(b.Length * i, SeekOrigin.Begin);
                fs.Write(b, 0, b.Length);

            }

            fs.Close();
            sw.Stop();
            Console.WriteLine("Stop - " + sw.ElapsedMilliseconds);
            sw.Reset();

            Console.ReadKey();
        }

        internal static void AddNodes(String depth, CFStorage cfs)
        {

            Action<CFItem> va = delegate (CFItem target)
            {

                String temp = target.Name + (target is CFStorage ? "" : " (" + target.Size + " bytes )");

                //Stream

                Console.WriteLine(depth + temp);

                if (target is CFStorage)
                {  //Storage

                    String newDepth = depth + "    ";

                    //Recursion into the storage
                    AddNodes(newDepth, (CFStorage)target);

                }
            };

            //Visit NON-recursively (first level only)
            cfs.VisitEntries(va, false);
        }

        internal static void CreateFile(string fileName)
        {
            const int MAX_STREAM_COUNT = 5000;

            CompoundFile cf = new CompoundFile();
            for (int i = 0; i < MAX_STREAM_COUNT; i++)
            {
                cf.RootStorage.AddStream("Test" + i.ToString()).SetData(Test.Helpers.GetBuffer(300));
            }
            cf.Save(fileName);
            cf.Close();
        }
    }
}
