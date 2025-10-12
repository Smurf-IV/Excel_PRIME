using System;
using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;

using BenchmarkDotNet.Attributes;

using Microsoft.VSDiagnostics;

namespace ExcelPRIME.Bench;

// For more information on the VS BenchmarkDotNet Diagnosers see https://learn.microsoft.com/visualstudio/profiling/profiling-with-benchmark-dotnet
[ExcludeFromCodeCoverage]
[CPUUsageDiagnoser]
public class Benchmarks
{
    private SHA256 sha256 = SHA256.Create();
    private byte[] data;

    [GlobalSetup]
    public void Setup()
    {
        data = new byte[10000];
        new Random(42).NextBytes(data);
    }

    [Benchmark]
    public byte[] Sha256()
    {
        return sha256.ComputeHash(data);
    }
}
