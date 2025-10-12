using System.Diagnostics.CodeAnalysis;

using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Running;

namespace ExcelPRIME.Bench;

[ExcludeFromCodeCoverage]
internal class Program
{
    static void Main(string[] args)
    {
        Summary[] _ = BenchmarkRunner.Run(typeof(Program).Assembly);
    }
}
