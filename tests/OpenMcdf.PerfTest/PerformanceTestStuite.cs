using System;
using System.Collections;
using System.Linq;
using NBench.Reporting.Targets;
using NBench.Sdk;
using NBench.Sdk.Compiler;
using NUnit.Framework;

namespace OpenMcdf.PerfTest
{
    public abstract class PerformanceTestStuite<T>
    {
        [TestCaseSource(nameof(Benchmarks))]
        public void PerformanceTests(Benchmark benchmark)
        {
            Benchmark.PrepareForRun();
            benchmark.Run();
            benchmark.Finish();
        }

        public static IEnumerable Benchmarks()
        {
            var discovery = new ReflectionDiscovery(new ActionBenchmarkOutput(report => { }, results =>
            {
                foreach (var assertion in results.AssertionResults)
                {
                    Assert.True(assertion.Passed, results.BenchmarkName + " " + assertion.Message);
                    Console.WriteLine(assertion.Message);
                }
            }));

            var benchmarks = discovery.FindBenchmarks(typeof(T)).ToList();

            foreach (var benchmark in benchmarks)
            {
                var name = benchmark.BenchmarkName.Split('+')[1];
                yield return new TestCaseData(benchmark).SetName(name);
            }
        }
    }
}