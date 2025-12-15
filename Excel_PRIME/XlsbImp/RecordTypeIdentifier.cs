namespace ExcelPRIME.XlsbImp;

/// <summary>
/// Creating an exhaustive list of all possible record type identifiers for .xlsb files would require referencing the official
/// Microsoft Excel Binary File Format (.xlsb) specification. Below is an expanded version of the RecordTypeIdentifier enum,
/// including more known record types. However, this list may still not be fully exhaustive, as the complete specification is extensive and proprietary.
/// </summary>
internal enum RecordTypeIdentifier
{
    // Workbook and Worksheet Records
    BOF = 0x0809,           // Beginning of File
    EOF = 0x000A,           // End of File
    BOOKBEGIN = 0x0083,     // 131: Start of WorkBook
    BOOKEND = 0x0084,       // 132: End of WorkBook
    SHEET = 0x0085,         // 133: Worksheet
    BUNDLESHEET = 0x009c,   // 156: Bundle Sheet
    DIMENSIONS = 0x0200,    // 148: Dimensions (used range of rows and columns)
    WINDOW1 = 0x003D,       // Window Information
    WINDOW2 = 0x023E,       // Sheet Window Information
    DATEMODE = 0x0022,      // Date Mode (1900 or 1904)
    CODEPAGE = 0x0042,      // Code Page
    BACKUP = 0x0040,        // Backup Flag
    FILESHARING = 0x005B,   // File Sharing Information
    WRITEACCESS = 0x005C,   // Write Access User Name
    FILEPASS = 0x002F,      // File Encryption Information
    PROTECT = 0x0012,       // Sheet Protection
    PASSWORD = 0x0013,      // Password Protection
    SCENPROTECT = 0x00DD,   // Scenario Protection
    OBJECTPROTECT = 0x0063, // Object Protection
    // Cell Records
    BLANK = 0x0201,         // Blank Cell
    INTEGER = 0x027E,       // Integer Cell
    NUMBER = 0x0203,        // Number Cell
    LABEL = 0x0204,         // Label Cell
    STRING = 0x0207,        // String Cell
    FORMULA = 0x0006,       // Formula Cell
    BOOLERR = 0x0205,       // Boolean or Error Cell
    // Formatting Records
    FORMAT = 0x041E,        // Format Record
    XF = 0x00E0,            // Extended Format
    STYLE = 0x0293,         // Style
    FONT = 0x0031,          // Font
    PALETTE = 0x0092,       // Color Palette
    THEME = 0x00A1,         // Theme Information
    // Shared Data Records
    STRINGITEM = 0x0013,    // Shared String
    SSTBEGIN = 0x009f,      // 159: Start of Shared String Table
    SSTEND = 0x00A0,        // 160: End of Shared String Table
    SHAREDSTRINGS = 0x00FC, // Shared Strings
    SST = 0x00FC,           // Shared String Table
    CONTINUE = 0x003C,      // Continuation Record
    // Row and Column Records
    ROW = 0x0208,           // Row Information
    COLINFO = 0x007D,       // Column Information
    MERGEDCELLS = 0x00E5,   // Merged Cells
    HIDDENCOL = 0x0081,     // Hidden Column
    HIDDENROW = 0x0082,     // Hidden Row
    // Chart and Drawing Records
    CHART = 0x0850,         // Chart
    DRAWING = 0x00EC,       // Drawing Object
    OBJ = 0x005D,           // Object (e.g., buttons, shapes)
    TXO = 0x01B6,           // Text Object
    IMDATA = 0x007F,        // Image Data
    // Hyperlink Records
    HYPERLINK = 0x01B8,     // Hyperlink
    // Print and Page Setup Records
    PRINTHEADERS = 0x002A,  // Print Headers
    PRINTGRIDLINES = 0x002B,// Print Gridlines
    HEADER = 0x0014,        // Header Text
    FOOTER = 0x0015,        // Footer Text
    PLS = 0x004D,           // Page Layout Settings
    SETUP = 0x00A1,         // Page Setup
    // Miscellaneous Records
    EXTERNNAME = 0x0023,    // External Name
    EXTERNSHEET = 0x0017,   // External Sheet
    NAME = 0x0018,          // Defined Name
    INDEX = 0x020B,         // Index Record
    NOTE = 0x001C,          // Comment or Note
    FILELOCK = 0x01A9,      // File Lock
    RECALCID = 0x01C1,      // Recalculation ID
    SHEETLAYOUT = 0x0862,   // Sheet Layout
    SHEETPROTECTION = 0x0867, // Sheet Protection Options
    CONDITIONALFORMATTING = 0x01B0, // Conditional Formatting
    AUTOFILTER = 0x009E,    // AutoFilter Information
    SORT = 0x0090,          // Sort Information
    QUERYTABLE = 0x0800,    // Query Table Information
    PIVOTTABLE = 0x00B9     // Pivot Table Information
}

