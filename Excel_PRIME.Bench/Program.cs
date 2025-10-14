using System.Diagnostics.CodeAnalysis;
using System.Linq;

using BenchmarkDotNet.Analysers;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Running;

namespace ExcelPRIME.Bench;

[ExcludeFromCodeCoverage]
internal class Program
{
    static void Main(string[] args)
    {
        Summary[] _ = BenchmarkRunner.Run(typeof(Program).Assembly, new Config());
    }
    private class Config : ManualConfig
    {
        public Config()
        {
            // DefaultJob
            //AddJob(Job.Dry);      // Dry(IterationCount=1, LaunchCount=1, RunStrategy=ColdStart, UnrollFactor=1, WarmupCount=1)
            AddJob(Job.ShortRun);     // ShortRun(IterationCount=3, LaunchCount=1, WarmupCount=3)
            //AddJob(Job.MediumRun);      // MediumRun(IterationCount=15, LaunchCount=2, WarmupCount=10)
            AddLogger(DefaultConfig.Instance.GetLoggers().ToArray()); // manual config has no loggers by default
            AddExporter(DefaultConfig.Instance.GetExporters().ToArray()); // manual config has no exporters by default
            AddColumnProvider(DefaultConfig.Instance.GetColumnProviders().ToArray()); // manual config has no columns by default
            //AddColumn(TargetMethodColumn.Method, StatisticColumn.Max);
            //AddExporter(HtmlExporter.Default, MarkdownExporter.GitHub, MarkdownExporter.Console);
            AddAnalyser(EnvironmentAnalyser.Default);
            AddDiagnoser(MemoryDiagnoser.Default);
            UnionRule = ConfigUnionRule.AlwaysUseLocal;
            SummaryStyle = SummaryStyle.Default.WithRatioStyle(RatioStyle.Trend);
        }
    }
}
