using System.Diagnostics.CodeAnalysis;
using System.Linq;

using BenchmarkDotNet.Analysers;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Environments;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Running;

namespace ExcelPRIME.Bench;

[ExcludeFromCodeCoverage]
internal class Program
{
    static void Main(string[] args)
    {
        string[] frameworks = ["net8.0", "net9.0", "net10.0"];

        foreach (var framework in frameworks)
        {
            BenchmarkRunner.Run(typeof(Program).Assembly, new Config(framework));
        }
    }
    private class MyRule : IColumnHidingRule
    {
        public bool NeedToHide(IColumn column) =>
            column.Category == ColumnCategory.Statistics
            || column.ColumnName == "RatioSD";
    }
    private class Config : ManualConfig
    {
        public Config(string framework)
        {
            // DefaultJob
            //AddJob(Job.Dry);      // Dry(IterationCount=1, LaunchCount=1, RunStrategy=ColdStart, UnrollFactor=1, WarmupCount=1)
            Runtime runtime = framework switch
            {
                "net8.0" => CoreRuntime.Core80,
                "net9.0" => CoreRuntime.Core90,
                "net10.0" => CoreRuntime.Core10_0,
                _ => CoreRuntime.Core80
            };
            AddJob(Job.ShortRun.WithRuntime(runtime).WithId(framework));     // ShortRun(IterationCount=3, LaunchCount=1, WarmupCount=3)
            //AddJob(Job.MediumRun);      // MediumRun(IterationCount=15, LaunchCount=2, WarmupCount=10)
            AddLogger(DefaultConfig.Instance.GetLoggers().ToArray()); // manual config has no loggers by default
            AddExporter(DefaultConfig.Instance.GetExporters().ToArray()); // manual config has no exporters by default
            AddColumnProvider(DefaultConfig.Instance.GetColumnProviders().ToArray()); // manual config has no columns by default
            //AddColumn(TargetMethodColumn.Method, StatisticColumn.Max);
            //AddExporter(HtmlExporter.Default, MarkdownExporter.GitHub, MarkdownExporter.Console);
            HideColumns(new MyRule());
            AddAnalyser(EnvironmentAnalyser.Default);
            AddDiagnoser(MemoryDiagnoser.Default);
            UnionRule = ConfigUnionRule.AlwaysUseLocal;
            SummaryStyle = SummaryStyle.Default.WithRatioStyle(RatioStyle.Trend);
        }
    }
}
