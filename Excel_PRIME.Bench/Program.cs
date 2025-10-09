using System.Diagnostics.CodeAnalysis;

using BenchmarkDotNet.Running;

namespace Excel_PRIME.Bench;

[ExcludeFromCodeCoverage]
internal class Program
{
    static void Main(string[] args)
    {
        var _ = BenchmarkRunner.Run(typeof(Program).Assembly);
    }
}
