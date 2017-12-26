using System;
using System.Diagnostics;
using System.IO;
using NBench;

namespace OpenMcdf.PerfTest
{
    public class PerfTest : PerformanceTestStuite<PerfTest>
    {
        private string _fileName;
        private Counter _testCounter;

        [PerfSetup]
        public void Setup(BenchmarkContext context)
        {
            _fileName = Path.Combine(Path.GetTempPath(), "PerfLoad.cfs");
            _testCounter = context.GetCounter("TestCounter");
        }

        [PerfBenchmark(Description = "Getting a stream of large file must take less than 200 ms",
            NumberOfIterations = 1, RunMode = RunMode.Iterations, TestMode = TestMode.Test, SkipWarmups = true)]
        [CounterTotalAssertion("TestCounter", MustBe.LessThanOrEqualTo, 200.0d)] // max 0.2 sec
        [CounterMeasurement("TestCounter")]
        public void Perf_SteamOfLargeFile()
        {
            if (!File.Exists(_fileName))
            {
                Helpers.CreateFile(_fileName);
            }

            using (var cf = new CompoundFile(_fileName))
            {
                Stopwatch sw = new Stopwatch();
                sw.Start();
                CFStream s = cf.RootStorage.GetStream("Test1");
                sw.Stop();

                var executionTime = sw.ElapsedMilliseconds;
                for (int i = 0; i < sw.ElapsedMilliseconds; i++)
                {
                    _testCounter.Increment();
                }

                Console.WriteLine($"Took {executionTime} seconds");
            }
        }

        [PerfCleanup]
        public void Cleanup()
        {
            if (File.Exists(_fileName))
            {
                File.Delete(_fileName);
            }
        }
    }
}
