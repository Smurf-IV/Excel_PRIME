namespace ExcelPRIME.XlsbImp;

/// <summary>
/// Creating an exhaustive list of all possible record type identifiers for .xlsb files would require referencing the official
/// Microsoft Excel Binary File Format (.xlsb) specification. Below is an expanded version of the RecordTypeIdentifier enum,
/// including more known record types. However, this list may still not be fully exhaustive, as the complete specification is extensive and proprietary.
/// </summary>
internal enum RecordTypeIdentifier
{
    // Workbook and Worksheet Records
    EOF = -1,           // End of File
    BOF = 0x0809,           // Beginning of File
    BOOKBEGIN = 0x0083,     // 131: Start of WorkBook
    BOOKEND = 0x0084,       // 132: End of WorkBook
    WSVIEWSSTART = 0x0085,  // 133: Worksheet Start
    WSVIEWSEND = 0x0086,    // 134: Worksheet End
    WSVIEWSTART = 0x0089,   // 137: WorkSheet View Start
    WSVIEWEND = 0x008A,     // 138: WorkSheet View End
    SHEETDATABEGIN = 0x0091,// 145: Begin Sheet Data
    SHEETDATAEND = 0x0092,  // 146: End Sheet Data
    SHEETPR = 0x0093,       // 147: 
    DIMENSION = 0x0094,     // 148: Dimension
    SELECTION = 0x0098,     // 152: 
    BUNDLESHEET = 0x009c,   // 156: Bundle Sheet
    DIMENSIONS = 0x0200,    // : Dimensions (used range of rows and columns)
    WINDOW1 = 0x003D,       // Window Information
    WINDOW2 = 0x023E,       // Sheet Window Information
    DATEMODE = 0x0022,      // Date Mode (1900 or 1904)
    CODEPAGE = 0x0042,      // Code Page
    BACKUP = 0x0040,        // Backup Flag
    FILESHARING = 0x005B,   // File Sharing Information
    WRITEACCESS = 0x005C,   // Write Access User Name
    FILEPASS = 0x002F,      // File Encryption Information
    PROTECT = 0x0012,       // Sheet Protection
    SCENPROTECT = 0x00DD,   // Scenario Protection
    OBJECTPROTECT = 0x0063, // Object Protection
    COLUMNSBEGIN = 0x186,   // 390
    SHEETFORMATPR = 0x01E5, // 485: Sheet Format Pr
    // Cell Records
    BLANK = 0x0201,         // Blank Cell
    INTEGER = 0x027E,       // Integer Cell
    NUMBER = 0x0203,        // Number Cell
    LABEL = 0x0204,         // Label Cell
    STRING = 0x0207,        // String Cell
    FORMULA = 0x0006,       // Formula Cell
    BOOLERR = 0x0205,       // Boolean or Error Cell
    // Shared Data Records
    STRINGITEM = 0x0013,    // Shared String
    SSTBEGIN = 0x009f,      // 159: Start of Shared String Table
    SSTEND = 0x00A0,        // 160: End of Shared String Table
    SST = 0x00FC,           // Shared String Table
    CONTINUE = 0x003C,      // Continuation Record
    // Row and Column Records
    BRTRWDESCENT = 0x0400,  // 1024: specifies the vertical distance in pixels from the bottom of the cell to the typographical baseline of the cell contents for the current row
    ROWHDR = 0x0000,
    DATAEND = 0x0092,       // 146: End of row data
    CELLBLANK = 0x0001,
    CELLRK = 0x0002,
    CELLERROR = 0x0003,
    CELLBOOL = 0x0004,
    CELLREAL = 0x0005,
    CELLST = 0x0006,
    CELLISST = 0x0007,
    CELLFMLASTRING = 0x0008,
    CELLFMLANUM = 0x0009,
    CELLFMLABOOL = 0x000A,
    CELLFMLAERROR = 0x000B,
    NOTE = 0x001C,          // Comment or Note

    ROWINFO = 0x0208,       // Row Information
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
    HEADER = 0x0014,        // Header Text
    FOOTER = 0x0015,        // Footer Text
    PRINTWIDTH = 0x0026,    // 38: width
    PRINTHEADERS = 0x002A,  // Print Headers
    PRINTGRIDLINES = 0x002B,// Print Gridlines
    PLS = 0x004D,           // Page Layout Settings
    SETUP = 0x00A1,         // Page Setup
    // Miscellaneous Records
    EXTERNNAME = 0x0023,    // External Name
    EXTERNSHEET = 0x0017,   // External Sheet
    NAME = 0x0018,          // Defined Name
    INDEX = 0x020B,         // Index Record
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

