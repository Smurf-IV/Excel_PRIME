using System;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelPRIME.Shared;

/// <summary>
/// Windows does not automatically delete temp files unless the user schedules a cleanup job.
/// However, you can set the file to be automatically deleted once your FileStream object goes out of scope.
/// You can do this by using the FileOptions.DeleteOnClose flag
/// </summary>
public sealed class TempFile : IDisposable
{
    private readonly FileInfo _fi;

    /// <summary>
    /// 
    /// </summary>
    /// <param name="extension">Optional extension of the temp file (Include the dot)</param>
    public TempFile(string? extension = null)
    {
        try
        {
            // Get the full name of the newly created Temporary file. 
            // Note that the GetTempFileName() method actually creates
            // a 0-byte file and returns the name of the created file.
            var fileName = System.IO.Path.GetTempFileName();
            // Create a FileInfo object to set the file's attributes
            _fi = new FileInfo(fileName);

            if (!string.IsNullOrWhiteSpace(extension))
            {
                FileAttributes attrs = _fi.Attributes;
                _fi.Delete();
                fileName += extension;
                _fi = new FileInfo(fileName);
                using (FileStream str = _fi.Create())
                {
                    // Empty
                }
                // Reset the attributes
                _fi.Attributes = attrs;
            }
            // Set the Attribute property of this file to Temporary. 
            // Although this is not completely necessary, the .NET Framework is able 
            // to optimize the use of Temporary files by keeping them cached in memory.
            _fi.Attributes |= FileAttributes.Temporary | FileAttributes.NotContentIndexed | FileAttributes.NoScrubData;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unable to create TEMP file or set its attributes: " + ex.Message);
        }
    }

    public FileStream OpenForAsyncWrite() => new FileStream(_fi.FullName, FileMode.Open, FileAccess.Write, FileShare.None, 32*1024, true);

    public FileStream OpenForAsyncRead() => new FileStream(_fi.FullName, FileMode.Open, FileAccess.Read, FileShare.None, 32 * 1024, true);


    // ReSharper disable UnusedMember.Local
    [Flags]
    private enum MoveFileFlags
    {
        None = 0,
        ReplaceExisting = 1,
        CopyAllowed = 2,
        DelayUntilReboot = 4,
        WriteThrough = 8,
        CreateHardlink = 16,
        FailIfNotTrackable = 32
    }
    // ReSharper restore UnusedMember.Local

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    private static extern bool MoveFileEx(string lpExistingFileName, string? lpNewFileName, MoveFileFlags dwFlags);

    ~TempFile()
    {
        Dispose(false);
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    private void Dispose(bool _)
    {
        if (_fi.Exists)
        {
            try
            {
                _fi.Delete();
            }
            catch
            {
                // best effort
                _fi.Attributes &= ~FileAttributes.Temporary;
                MoveFileEx(_fi.FullName, null, MoveFileFlags.DelayUntilReboot);
            }
        }
    }
}
