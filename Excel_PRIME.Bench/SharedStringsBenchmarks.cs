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
#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
    public string FileName { get; set; }
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.

    private ZipArchive? _archive;
    private Stream? _sharedStringsStream;
    private ISharedString? _sharedStrings;

    [GlobalSetup]
    public void Setup()
    {
        string path = Path.Combine(RootFolder, FileName);
        FileStream fs = File.OpenRead(path);
        _archive = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);
        ZipArchiveEntry? entry = _archive.GetEntry("xl/sharedStrings.xml");
        if (entry == null)
        {
            return;
        }
        _sharedStringsStream = entry.Open();

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
                    Task<ISharedString> task = (Task<ISharedString>)getSharedStringsAsync.Invoke(helper, [_sharedStringsStream, CancellationToken.None])!;
                    task.Wait();
                    _sharedStrings = task.Result;
                }
            }
        }
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        _sharedStrings?.Dispose();
        _sharedStringsStream?.Dispose();
        _archive?.Dispose();
    }

    //[Benchmark(Baseline = true)]
    public int AccessFirstThousandSequential()
    {
        if (_sharedStrings is null)
        {
            return 0;
        }

        int total = 0;
        for (int i = 0; i < 1000; i++)
        {
            string? s = _sharedStrings[i];
            if (s != null)
            {
                total += s.Length;
            }
        }
        return total;
    }

    //[Benchmark]
    public int AccessRandomThousand()
    {
        if (_sharedStrings is null)
        {
            return 0;
        }

        int total = 0;
        Random rnd = new Random(42);
        for (int i = 0; i < 1000; i++)
        {
            int idx = rnd.Next(0, 5000);
            string? s = _sharedStrings[idx];
            if (s != null)
            {
                total += s.Length;
            }
        }
        return total;
    }
}
