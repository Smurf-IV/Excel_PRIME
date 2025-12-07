using System;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

using BenchmarkDotNet.Attributes;

namespace ExcelPRIME.Bench;

[MemoryDiagnoser]
public class SharedStringsBenchmarks
{
    private const string RootFolder = "Data";

    [Params(
        "100mb.xlsx",
        "sampledocs-50mb-xlsx-file-sst.xlsx"
    )]
    public string FileName { get; set; }

    private ZipArchive? archive;
    private Stream? sharedStringsStream;
    private ISharedString? sharedStrings;

    [GlobalSetup]
    public void Setup()
    {
        string path = Path.Combine(RootFolder, FileName);
        FileStream fs = File.OpenRead(path);
        archive = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);
        ZipArchiveEntry? entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry == null)
        {
            return;
        }
        sharedStringsStream = entry.Open();

        // Instantiate internal XmlReaderHelpersAsync via reflection and call GetSharedStringsAsync
        Assembly asm = Assembly.Load("Excel_PRIME");
        Type? helperType = asm.GetType("ExcelPRIME.Implementation.XmlReaderHelpersAsync", throwOnError: false, ignoreCase: false);
        if (helperType != null)
        {
            // Create instance (internal) via Activator
            object? helper = Activator.CreateInstance(helperType, nonPublic: true);
            if (helper != null)
            {
                MethodInfo? getSharedStringsAsync = helperType.GetMethod("GetSharedStringsAsync", BindingFlags.Public | BindingFlags.Instance | BindingFlags.NonPublic);
                if (getSharedStringsAsync != null)
                {
                    Task<ISharedString> task = (System.Threading.Tasks.Task<ISharedString>)getSharedStringsAsync.Invoke(helper, new object[] { sharedStringsStream, CancellationToken.None })!;
                    task.Wait();
                    sharedStrings = task.Result;
                }
            }
        }
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        sharedStrings?.Dispose();
        sharedStringsStream?.Dispose();
        archive?.Dispose();
    }

    [Benchmark(Baseline = true)]
    public int AccessFirstThousandSequential()
    {
        if (sharedStrings is null)
        {
            return 0;
        }

        int total = 0;
        for (int i = 0; i < 1000; i++)
        {
            string? s = sharedStrings[i];
            if (s != null)
            {
                total += s.Length;
            }
        }
        return total;
    }

    [Benchmark]
    public int AccessRandomThousand()
    {
        if (sharedStrings is null)
        {
            return 0;
        }

        int total = 0;
        Random rnd = new Random(42);
        for (int i = 0; i < 1000; i++)
        {
            int idx = rnd.Next(0, 5000);
            string? s = sharedStrings[idx];
            if (s != null)
            {
                total += s.Length;
            }
        }
        return total;
    }
}
