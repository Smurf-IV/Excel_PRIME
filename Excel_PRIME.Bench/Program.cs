using System.Diagnostics.CodeAnalysis;

using BenchmarkDotNet.Analysers;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Exporters;
using BenchmarkDotNet.Exporters.Csv;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Loggers;
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
            AddJob(Job.Dry);
            AddLogger(ConsoleLogger.Default);
            AddColumn(TargetMethodColumn.Method, StatisticColumn.Max);
            AddExporter(HtmlExporter.Default, MarkdownExporter.GitHub, MarkdownExporter.Console);
            AddAnalyser(EnvironmentAnalyser.Default);
            AddDiagnoser(MemoryDiagnoser.Default);
            UnionRule = ConfigUnionRule.AlwaysUseLocal;
            SummaryStyle = SummaryStyle.Default.WithRatioStyle(RatioStyle.Trend);
        }
    }
}
