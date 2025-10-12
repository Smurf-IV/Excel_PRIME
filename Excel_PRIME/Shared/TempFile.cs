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
    public FileInfo FileInfo { get; private set; } = null!;

    /// <summary>
    /// 
    /// </summary>
    /// <param name="extension">Optional extension of the temp file (Include the dot)</param>
    public TempFile(string? extension = null)
    {
        CreateTempFile(extension);
    }

    private TempFile(bool _)
    {
    }

    /// <summary>
    /// Creates the file (If necesary), and sets attributes to Temp
    /// </summary>
    public static TempFile MakeThisATempFile(string fullName)
    {
        TempFile tempFile = new TempFile(false);
        FileInfo fileInfo = new FileInfo(fullName);
        if (!fileInfo.Exists)
        {
            DirectoryInfo? dirInfo = fileInfo.Directory;
            dirInfo?.Create();

            using FileStream str = fileInfo.Create();
        }
        tempFile.MakeThisATempFileImpl(fileInfo);
        return tempFile;
    }

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
        if (FileInfo!.Exists)
        {
            try
            {
                FileInfo.Delete();
            }
            catch
            {
                // best effort
                FileInfo.Attributes &= ~FileAttributes.Temporary;
                MoveFileEx(FileInfo.FullName, null, MoveFileFlags.DelayUntilReboot);
            }
        }
    }

    private string CreateTempFile(string? extension)
    {
        string fileName = string.Empty;

        try
        {
            // Get the full name of the newly created Temporary file. 
            // Note that the GetTempFileName() method actually creates
            // a 0-byte file and returns the name of the created file.
            fileName = System.IO.Path.GetTempFileName();
            // Create a FileInfo object to set the file's attributes
            FileInfo fileInfo = new FileInfo(fileName);

            if (!string.IsNullOrWhiteSpace(extension))
            {
                FileAttributes attrs = fileInfo.Attributes;
                fileInfo.Delete();
                fileName += extension;
                fileInfo = new FileInfo(fileName);
                using (FileStream str = fileInfo.Create())
                {
                    // Empty
                }
                // Reset the attributes
                fileInfo.Attributes = attrs;
            }
            MakeThisATempFileImpl(fileInfo);

            Console.WriteLine("TEMP file created at: " + fileName);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unable to create TEMP file or set its attributes: " + ex.Message);
        }

        return fileName;
    }

    private void MakeThisATempFileImpl(FileInfo fileInfo)
    {
        FileInfo = fileInfo;

        // Set the Attribute property of this file to Temporary. 
        // Although this is not completely necessary, the .NET Framework is able 
        // to optimize the use of Temporary files by keeping them cached in memory.
        FileInfo.Attributes |= FileAttributes.Temporary;
    }

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

}
